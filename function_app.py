import logging
import sys
import os
import json
import requests
import azure.functions as func
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient
from datetime import datetime
from urllib.parse import quote, urljoin
import base64
import re
import traceback
from typing import List, Dict, Optional

# ---------- Logging & console encoding ----------
logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
try:
    # Ensure Windows console handles Unicode during local dev
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _preview(obj, limit=1200) -> str:
    """Safe, short preview for logs (escapes non-ASCII)."""
    try:
        s = json.dumps(obj, ensure_ascii=True)
    except Exception:
        s = str(obj)
    return s[:limit] + ("..." if len(s) > limit else "")


# ---------- Azure Functions app ----------
app = func.FunctionApp()

REQUIRED_ENV_VARS = [
    "AI_FOUND_ENDPOINT",
    "AI_FOUND_API_KEY",
    "SEARCH_ENDPOINT",
    "SEARCH_INDEX_NAME",
    "SEARCH_KEY",
]


def get_env_var(key, default=None):
    val = os.environ.get(key, default)
    if val is None:
        raise EnvironmentError(f"Missing required environment variable: {key}")
    return val

@app.function_name(name="TimerTrigger")
@app.schedule(schedule="0 */30 * * * *", arg_name="mytimer", run_on_startup=True, use_monitor=True)
def main(mytimer: func.TimerRequest) -> None:
    logging.info("🔁 TimerTrigger sync started.")

    try:
        TOTAL_CAP = int(os.environ.get("MAX_FILES", "250"))
        processed_files = 0
        BLOB_PREFIX = "sharepoint-data"
        PROGRESS_FILE = f"{BLOB_PREFIX}/.progress.json"

        blob_conn = os.environ["BLOB_CONNECTION_STRING"]
        container = os.environ.get("SHAREPOINT_CONTAINER", "your-container-name")
        hostname = os.environ["TENANT_NAME"]
        sharepoint_base_url = os.environ.get("SHAREPOINT_BASE_URL", "").rstrip("/")

        drive_targets = json.loads(os.environ["SHAREPOINT_DRIVES"])

        blob_service = BlobServiceClient.from_connection_string(blob_conn)
        blob_client = blob_service.get_blob_client(container=container, blob=PROGRESS_FILE)

        credential = DefaultAzureCredential()
        token = credential.get_token("https://graph.microsoft.com/.default").token
        headers = {"Authorization": f"Bearer {token}"}

        # Load previous progress (may be empty)
        try:
            progress = json.loads(blob_client.download_blob().readall())
        except Exception:
            progress = {}

        new_progress = json.loads(json.dumps(progress))

        def save_progress():
            _bc = blob_service.get_blob_client(container=container, blob=PROGRESS_FILE)
            _bc.upload_blob(json.dumps(new_progress, indent=2), overwrite=True)

        def resolve_site_id(site_path: str) -> str:
            resp = requests.get(f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}", headers=headers)
            resp.raise_for_status()
            return resp.json()["id"]

        def resolve_drive_id(site_id: str, drive_name: str) -> str:
            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
            resp = requests.get(url, headers=headers)
            resp.raise_for_status()
            for d in resp.json().get("value", []):
                if d["name"].lower() == drive_name.lower():
                    return d["id"]
            raise Exception(f"Drive '{drive_name}' not found")

        def get_item_language(drive_id: str, item_id: str) -> Optional[str]:
            """
            Fetch the 'Language' column from the listItem fields for a given drive item.
            Adjust 'Language' below if your internal column name is different.
            """
            try:
                url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/listItem?$expand=fields"
                resp = requests.get(url, headers=headers)
                resp.raise_for_status()
                data = resp.json()
                fields = data.get("fields", {})

                # IMPORTANT: 'Language' must match the internal name of your column
                return fields.get("Language")
            except Exception as e:
                logging.warning("⚠️ Failed to read Language field for item %s: %s", item_id, e)
                return None

        def list_all_files(drive_id: str, folder_id: str = "root", folder_path: str = ""):
            items = []
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children"
            while url:
                resp = requests.get(url, headers=headers)
                resp.raise_for_status()
                data = resp.json()

                for item in data.get("value", []):
                    name = item["name"]
                    path = f"{folder_path}/{name}".strip("/")
                    if "folder" in item:
                        items += list_all_files(drive_id, item["id"], path)
                    else:
                        # 🔹 NEW: fetch the Language field from listItem
                        language = get_item_language(drive_id, item["id"])

                        items.append({
                            "id": item["id"],
                            "name": name,
                            "path": path,
                            "lastModifiedDateTime": item["lastModifiedDateTime"],
                            "webUrl": item.get("webUrl"),
                            "size": item.get("size"),
                            "file": item.get("file", {}),
                            "language": language,  # 🔹 NEW
                        })
                url = data.get("@odata.nextLink")
            return items

        def upload_file(blob_path: str, content_url: str, site_name: str, drive_name: str, language: Optional[str]):
            nonlocal processed_files
            with requests.get(content_url, headers=headers, stream=True) as r:
                if r.status_code == 200:
                    b = blob_service.get_blob_client(container=container, blob=blob_path)
                    meta = {
                        "site": site_name,
                        "directory": drive_name
                    }
                    if language:
                        meta["language"] = language  # 🔹 NEW
                    # Try to capture a resolvable SP URL if available
                    doc_url = getattr(r, "url", None)
                    if doc_url:
                        meta["url"] = doc_url
                    b.upload_blob(r.raw, overwrite=True, metadata=meta)
                    processed_files += 1
                    logging.info("📤 Uploaded: %s (dir=%s)", blob_path, drive_name)
                    return True
                else:
                    logging.warning("⚠️ Download failed %s (%s)", content_url, r.status_code)
                    return False

        def upload_metadata(blob_path: str, metadata: dict):
            # Always include valid SharePoint URL in metadata as 'url'
            meta = dict(metadata)
            doc_url = meta.get("webUrl")
            if not doc_url and sharepoint_base_url:
                file_path = meta.get("path")
                if file_path:
                    doc_url = urljoin(sharepoint_base_url + "/", quote(file_path, safe="/"))
            if doc_url:
                meta["url"] = doc_url
            b = blob_service.get_blob_client(container=container, blob=blob_path + ".metadata.json")
            b.upload_blob(json.dumps(meta), overwrite=True)

        def get_parent_id(blob_path: str) -> Optional[str]:
            """
            Given either:
            - a content blob path: .../file.ext
            - OR its metadata blob path: .../file.ext.metadata.json

            Read the metadata JSON and compute the parent_id exactly
            as stored in the search index: base64(url).
            """
            try:
                # Accept both content and metadata paths
                if blob_path.endswith(".metadata.json"):
                    meta_blob_name = blob_path
                else:
                    meta_blob_name = blob_path + ".metadata.json"

                meta_blob = blob_service.get_blob_client(
                    container=container,
                    blob=meta_blob_name
                )
                data = json.loads(meta_blob.download_blob().readall())

                url = data.get("url")
                if not url:
                    logging.warning("⚠️ No 'url' field in metadata for %s; cannot compute parent_id", blob_path)
                    return None

                parent_id = base64.b64encode(url.encode("utf-8")).decode("ascii")
                return parent_id

            except Exception as e:
                logging.warning("⚠️ Failed to compute parent_id for %s: %s", blob_path, e)
                return None

        def get_chunks_for_parent(parent_id: str) -> List[str]:
            search_endpoint = os.environ["SEARCH_ENDPOINT"].rstrip("/")
            index_name = os.environ["SEARCH_INDEX_NAME"]
            api_key = os.environ["SEARCH_KEY"]
            safe_parent = parent_id.replace("'", "''")

            url = f"{search_endpoint}/indexes/{index_name}/docs"
            params = {
                "api-version": "2024-07-01",
                "$select": "chunk_id",
                "$filter": f"parent_id eq '{safe_parent}'",
                "$top": 1000
            }

            logging.info("🔎 Querying index for parent_id=%s", parent_id)
            resp = requests.get(url, headers={"api-key": api_key}, params=params)

            if resp.status_code != 200:
                logging.error(
                    "❌ Failed to fetch chunks for parent_id=%s | status=%s body=%s",
                    parent_id, resp.status_code, resp.text
                )
                return []

            data = resp.json()
            values = data.get("value", [])
            chunk_ids = [d["chunk_id"] for d in values if "chunk_id" in d]

            logging.info(
                "📄 Found %d chunks for parent_id=%s",
                len(chunk_ids),
                parent_id
            )
            if not chunk_ids:
                logging.warning("⚠️ No chunks found to delete for parent_id=%s", parent_id)

            return chunk_ids

        def delete_from_search_index(doc_keys: List[str]):
            if not doc_keys:
                logging.info("ℹ️ No doc_keys passed to delete_from_search_index; skipping.")
                return

            search_endpoint = os.environ["SEARCH_ENDPOINT"].rstrip("/")
            index_name = os.environ["SEARCH_INDEX_NAME"]
            key = os.environ["SEARCH_KEY"]

            logging.info("🧹 Attempting to delete %d chunks from index '%s'", len(doc_keys), index_name)

            payload = {
                "value": [
                    {"@search.action": "delete", "chunk_id": k}
                    for k in doc_keys
                ]
            }

            resp = requests.post(
                f"{search_endpoint}/indexes/{index_name}/docs/index?api-version=2024-07-01",
                headers={"api-key": key, "Content-Type": "application/json"},
                json=payload
            )

            if resp.status_code != 200:
                logging.error(
                    "❌ Failed index delete (%d keys). status=%s body=%s",
                    len(doc_keys),
                    resp.status_code,
                    resp.text
                )
            else:
                try:
                    body = resp.json()
                except Exception:
                    body = resp.text[:500]

                logging.info(
                    "✅ Index delete request accepted for %d chunks. Response: %s",
                    len(doc_keys),
                    body
                )

        def delete_orphans(blob_paths_expected: set):
            container_client = blob_service.get_container_client(container)

            orphan_content_blobs: List[str] = []
            metadata_blobs: List[str] = []
            content_blobs_seen: set[str] = set()
            parent_ids_to_delete: set[str] = set()

            # 1) Scan all blobs once
            existing = container_client.list_blobs(name_starts_with=BLOB_PREFIX)
            for bl in existing:
                path = bl.name

                # Skip progress tracking file
                if path.endswith(".progress.json"):
                    continue

                # Collect metadata blobs for later
                if path.endswith(".metadata.json"):
                    metadata_blobs.append(path)
                    continue

                # Everything else is treated as a content blob
                content_blobs_seen.add(path)

                # Content blob is orphan if it's not expected this run
                if path not in blob_paths_expected:
                    orphan_content_blobs.append(path)
                    pid = get_parent_id(path)
                    if pid:
                        parent_ids_to_delete.add(pid)

            # 2) Identify metadata-only orphans (no content blob exists anymore)
            orphan_metadata_blobs: List[str] = []
            for meta_path in metadata_blobs:
                base_path = meta_path[: -len(".metadata.json")]

                # Only metadata orphan if:
                # - it's not expected this run
                # - we did NOT see a content blob for it (so we won't delete it as a sibling)
                if base_path not in blob_paths_expected and base_path not in content_blobs_seen:
                    orphan_metadata_blobs.append(meta_path)
                    pid = get_parent_id(meta_path)
                    if pid:
                        parent_ids_to_delete.add(pid)

            logging.info(
                "🧮 Orphan scan complete. orphan_content=%d orphan_metadata=%d unique_parent_ids=%d",
                len(orphan_content_blobs),
                len(orphan_metadata_blobs),
                len(parent_ids_to_delete),
            )

            # 3) Delete index documents for all collected parent_ids
            for pid in parent_ids_to_delete:
                logging.info("🔧 Processing index cleanup for parent_id=%s", pid)
                chunk_ids = get_chunks_for_parent(pid)
                delete_from_search_index(chunk_ids)

            # 4) Delete orphan content blobs and their metadata siblings
            for blob_path in orphan_content_blobs:
                try:
                    container_client.get_blob_client(blob=blob_path).delete_blob()
                    logging.info("🗑️ Deleted orphan blob: %s", blob_path)
                except Exception as e:
                    logging.exception("⚠️ Failed to delete orphan blob %s: %s", blob_path, e)

                # Best-effort delete metadata sibling
                meta_blob_path = blob_path + ".metadata.json"
                try:
                    container_client.get_blob_client(blob=meta_blob_path).delete_blob()
                    logging.info("🗑️ Deleted orphan metadata blob: %s", meta_blob_path)
                except Exception:
                    # If it's already gone, that's fine – don't log as error
                    pass

            # 5) Delete orphan metadata-only blobs (no content blob exists)
            for meta_path in orphan_metadata_blobs:
                try:
                    container_client.get_blob_client(blob=meta_path).delete_blob()
                    logging.info("🗑️ Deleted orphan metadata-only blob: %s", meta_path)
                except Exception as e:
                    logging.exception("⚠️ Failed to delete orphan metadata blob %s: %s", meta_path, e)

        # === Process drives ===
        blob_paths_to_keep = set()

        # ✅ Per-drive quota so one drive cannot starve others
        per_drive_cap = max(1, TOTAL_CAP // max(1, len(drive_targets)))

        any_drive_succeeded = False
        any_drive_failed = False

        for entry in drive_targets:
            site_path = entry["site_path"]
            site_name = entry.get("site_name", site_path.strip("/").split("/")[-1])
            drive_name = entry["drive_name"]
            key = f"{site_name}:{drive_name}"

            logging.info("🔄 Processing drive: %s", drive_name)
            prev = progress.get(key, {})
            new_progress.setdefault(key, dict(prev))

            processed_files_this_drive = 0

            try:
                site_id = resolve_site_id(site_path)
                drive_id = resolve_drive_id(site_id, drive_name)

                files = list_all_files(drive_id)
                any_drive_succeeded = True  # ✅ we got at least one drive listing
                seen_paths = set()

                for f in files:
                    # Skip large/media files
                    if f["name"].lower().endswith(".mp4") or (f.get("size") and f["size"] > 100 * 1024 * 1024):
                        logging.warning("⏩ Skipping large/media file: %s (%s bytes)", f["name"], f.get("size"))
                        continue

                    file_key = f["path"]
                    blob_path = f"{BLOB_PREFIX}/{site_name}/{drive_name}/{file_key}"
                    seen_paths.add(file_key)
                    blob_paths_to_keep.add(blob_path)

                    last_mod = f["lastModifiedDateTime"]
                    prev_mod = prev.get(file_key)
                    needs_upload = (prev_mod != last_mod)

                    # Upload (capped globally & per-drive)
                    if needs_upload and processed_files < TOTAL_CAP and processed_files_this_drive < per_drive_cap:
                        ok = upload_file(
                            blob_path,
                            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{f['id']}/content",
                            site_name,
                            drive_name,
                            f.get("language")  # 🔹 NEW
                        )
                        if ok:
                            upload_metadata(blob_path, f)
                            new_progress[key][file_key] = last_mod
                            processed_files_this_drive += 1
                        else:
                            logging.warning("⏸️ Skipped progress (download/upload failed): %s", blob_path)
                    else:
                        if not needs_upload and prev_mod:
                            new_progress[key][file_key] = prev_mod
                        elif needs_upload:
                            # We *wanted* to upload but hit caps – do NOT advance progress
                            logging.warning("⏸️ Upload cap reached; will upload next run: %s", blob_path)

                    # Backfill metadata if blob exists (doesn't imply progress)
                    b = blob_service.get_blob_client(container=container, blob=blob_path)
                    try:
                        props = b.get_blob_properties()
                        meta = props.metadata or {}
                        wanted = {"site": site_name, "directory": drive_name}
                        doc_url = f.get("webUrl")
                        if not doc_url and sharepoint_base_url:
                            file_path = f.get("path")
                            if file_path:
                                doc_url = urljoin(sharepoint_base_url + "/", quote(file_path, safe="/"))
                        if doc_url:
                            wanted["url"] = doc_url
                        language = f.get("language")
                        if language:
                            wanted["language"] = language  # 🔹 NEW
                        if any(meta.get(k) != v for k, v in wanted.items()):
                            meta.update(wanted)
                            b.set_blob_metadata(metadata=meta)
                            logging.info("📝 Backfilled metadata for: %s (dir=%s)", blob_path, drive_name)
                    except Exception:
                        pass

                logging.info(
                    "✅ Finished drive: %s | uploaded this drive=%d (per-drive cap=%d), total uploaded=%d/%d",
                    drive_name, processed_files_this_drive, per_drive_cap, processed_files, TOTAL_CAP
                )

                new_progress[key] = {k: v for k, v in new_progress[key].items() if k in seen_paths}

                # Save progress per drive to avoid losing work on long runs
                save_progress()

            except Exception as e:
                any_drive_failed = True
                logging.warning("⚠️ Skipping drive '%s': %s", drive_name, e)

        # Orphan cleanup + final save (SAFE)
        if not any_drive_succeeded:
            logging.error(
                "🚫 No drives processed successfully in this run. "
                "Skipping orphan cleanup to avoid accidental mass deletions."
            )
        elif any_drive_failed:
            logging.warning(
                "⚠️ Some drives failed during this run. "
                "Skipping orphan cleanup to avoid deleting still-valid blobs."
            )
        else:
            delete_orphans(blob_paths_to_keep)

        save_progress()
        logging.info("✅ Sync complete. Files uploaded this run: %d (cap=%d)", processed_files, TOTAL_CAP)

    except Exception:
        logging.exception("❌ TimerTrigger execution failed")

def load_system_prompt(path: str = "system_prompt.md") -> str:
    try:
        if not os.path.isabs(path):
            path = os.path.join(os.path.dirname(__file__), path)
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        logging.error("System prompt file not found at %s", path)
        return "System prompt file not found."

def decode_blob_url(chunk_or_parent_id: str) -> Optional[str]:
    try:
        # Extract the encoded part (usually after last underscore)
        encoded = chunk_or_parent_id.split("_")[-1] if "_" in chunk_or_parent_id else chunk_or_parent_id

        # Add padding for base64 decoding
        padded = encoded + "=" * ((4 - len(encoded) % 4) % 4)

        # Decode base64
        decoded = base64.b64decode(padded).decode("utf-8")

        # Remove blob storage prefix if present
        if decoded.startswith("sharepoint-data/"):
            decoded = decoded.replace("sharepoint-data/", "")

        # Remove site name + drive if present (keep only the actual file path)
        parts = decoded.split("/")
        if len(parts) > 2:
            decoded = "/".join(parts[2:])

        # URL encode the path (keep slashes and colon for safety)
        return quote(decoded, safe=":/")
    except Exception as e:
        logging.warning("Failed to decode blob URL from %s: %s", chunk_or_parent_id, e)
        return None


# --------- NEW HELPERS FOR DISPLAY NAMES (no URL construction) ----------
def normalize_url(u: str) -> str:
    return (u or "").lower().split("?", 1)[0].rstrip("/")

def _leaf(path: str) -> Optional[str]:
    if not path:
        return None
    leaf = path.strip("/").split("/")[-1]
    return leaf or None

def pick_display_name(cit_doc: Dict, index_doc: Optional[Dict], url: str) -> str:
    """
    Best-effort, URL-safe display label without constructing a new URL.
    Priority:
      1) citation doc title (if not *.aspx)
      2) index_doc.title (if present & not *.aspx)
      3) index_doc.key leaf name (actual file name from index path)
      4) URL query parameters (file/filename/originalPath/id)
      5) URL path tail if not *.aspx
    """
    def not_viewerish(name: Optional[str]) -> bool:
        if not name: return False
        n = name.strip().lower()
        return n and not n.endswith(".aspx")

    # 1) citation-provided title
    t = (cit_doc.get("title") or "").strip()
    if not_viewerish(t):
        return t

    # 2) index doc title
    if index_doc:
        t2 = (index_doc.get("title") or "").strip()
        if not_viewerish(t2):
            return t2
        # 3) key leaf (often holds true file name)
        k_leaf = _leaf(index_doc.get("key", ""))
        if not_viewerish(k_leaf):
            return k_leaf

    # 4) try URL query params / 5) last path segment
    try:
        from urllib.parse import urlparse, parse_qs, unquote
        p = urlparse(url)
        qs = parse_qs(p.query)
        for cand in ("file", "filename", "originalPath", "id"):
            if cand in qs and qs[cand]:
                tail = unquote(qs[cand][0]).split("/")[-1]
                if not_viewerish(tail):
                    return tail
        last = unquote(p.path.rstrip("/").split("/")[-1])
        if not_viewerish(last):
            return last
    except Exception:
        pass

    return "Document"


def remove_orphan_citations(text: str, valid_numbers: set) -> str:
    return re.sub(r"\[(\d+)\]", lambda m: f"[{m.group(1)}]" if m.group(1) in valid_numbers else "", text)

def _append_reference_links(answer: str, docs: List[Dict]) -> str:
    if not docs or not isinstance(docs, list):
        return answer

    # Fix joined citations like [1][2] and [1][2][3]
    answer = re.sub(r"\[(\d+)\]\[(\d+)\]", r"[\1], [\2]", answer)
    answer = re.sub(r"(\[\d+\])(?=\[\d+\])", r"\1, ", answer)

    # Extract used reference numbers
    used_refs = set(re.findall(r"\[(\d+)\]", answer))
    if not used_refs:
        return answer

    # Prepare references section
    references_section = "\n\n**Sources:**\n"
    base_url = os.environ.get("SHAREPOINT_BASE_URL", "").rstrip("/")

    for ref in sorted(used_refs, key=int):
        i = int(ref) - 1
        if 0 <= i < len(docs):
            # Only use url from index
            raw_url = (docs[i].get("url") or "").strip()
            if raw_url:
                references_section += f"[{ref}]: {raw_url}\n"
                logging.debug("🔗 Citation [%s] → %s", ref, raw_url)
            else:
                references_section += f"[{ref}]: #\n"
                logging.warning("⚠️ Missing url for citation [%s]", ref)
        else:
            references_section += f"[{ref}]: #\n"
            logging.warning("⚠️ Reference index [%s] out of bounds", ref)

    return answer.strip() + "\n\n" + references_section.strip() + "\n"

def _error_response(message: str, status_code: int, extra: Optional[dict] = None) -> func.HttpResponse:
    payload = {"error": message}
    if extra:
        payload.update(extra)
    # Print the exact error payload to console too
    logging.error("Sending error response: %s", _preview(payload))
    print("ERROR RESPONSE:", _preview(payload))
    return func.HttpResponse(json.dumps(payload), status_code=status_code, mimetype="application/json")

# ---------- HTTP Trigger ----------
@app.function_name(name="HttpAskAI")
@app.route(route="ask-ai", auth_level=func.AuthLevel.FUNCTION)
def ask_ai(req: func.HttpRequest) -> func.HttpResponse:
    try:
        # Body
        try:
            req_body = req.get_json()
        except ValueError:
            return _error_response("Request body must be valid JSON", 400)

        logging.info("Request payload: %s", _preview(req_body))

        message = req_body.get("message")
        chat_history = req_body.get("history", [])
        directories = req_body.get("directories", [])
        instructions = req_body.get("instructions", "")

        if not message:
            return _error_response("Missing 'message' in request body", 400)

        if not directories:
            payload = {
                "answer": "Please select at least one content directory before asking your question.",
                "references": []
            }
            logging.info("Sending early response (no directories): %s", _preview(payload))
            print("RESPONSE (no dirs):", _preview(payload))
            return func.HttpResponse(json.dumps(payload), status_code=200, mimetype="application/json")

        missing = [k for k in REQUIRED_ENV_VARS if not os.getenv(k)]
        if missing:
            return _error_response(f"Missing environment variable(s): {missing}", 500)

        # OData-safe filter (escape single quotes)
        safe_dirs = [str(d).replace("'", "''") for d in directories]
        filter_expr = " or ".join([f"directory eq '{d}'" for d in safe_dirs])
        logging.info("Azure Search filter expression: %s", filter_expr)

        # System prompt & messages
        system_prompt = load_system_prompt()
        if instructions:
            system_prompt += f"\n\nAdditional instructions: {instructions}\n"
        system_prompt += f"\nLimit your answer to the following content directories: {', '.join(directories)}."

        messages = [{"role": "system", "content": system_prompt}] + chat_history + [{"role": "user", "content": message}]
        logging.info(
            "Messages meta: %s",
            _preview({"dirs": directories, "system_len": len(system_prompt), "history_len": len(chat_history), "user_len": len(message or "")})
        )

        # Call AI Foundational endpoint
        ai_response = requests.post(
            os.environ["AI_FOUND_ENDPOINT"],
            headers={"api-key": os.environ["AI_FOUND_API_KEY"], "Content-Type": "application/json"},
            json={
                "messages": messages,
                "temperature": 0.2,
                "top_p": 1.0,
                "data_sources": [
                    {
                        "type": "azure_search",
                        "parameters": {
                            "endpoint": os.environ["SEARCH_ENDPOINT"],
                            "index_name": os.environ["SEARCH_INDEX_NAME"],
                            "semantic_configuration": "your-semantic-configuration",
                            "query_type": "semantic",
                            "fields_mapping": {},
                            "in_scope": True,
                            "filter": filter_expr,
                            "strictness": 3,
                            "top_n_documents": 5,
                            "authentication": {
                                "type": "api_key",
                                "key": os.environ["SEARCH_KEY"]
                            }
                        }
                    }
                ]
            },
            timeout=90
        )

        logging.info("AI service status code: %s", ai_response.status_code)
        logging.info("AI service response body (preview): %s", ai_response.text[:2000])

        if ai_response.status_code != 200:
            return _error_response("AI service request failed", ai_response.status_code, extra={"message": ai_response.text})

        result = ai_response.json()
        logging.info("AI service parsed keys: %s", list(result.keys()))

        # No choices
        if "choices" not in result or not result["choices"]:
            payload = {"answer": "No answer returned from AI service.", "references": []}
            logging.info("Sending response (no choices): %s", _preview(payload))
            print("RESPONSE (no choices):", _preview(payload))
            return func.HttpResponse(json.dumps(payload), status_code=200, mimetype="application/json")

        assistant_reply = result["choices"][0]["message"]["content"]
        citations = result["choices"][0]["message"].get("context", {}).get("citations", [])
        logging.info("Assistant reply (preview): %s", assistant_reply[:800])
        logging.info("Citations count: %s", len(citations))

        search_resp = requests.get(
            f"{os.environ['SEARCH_ENDPOINT'].rstrip('/')}/indexes/{os.environ['SEARCH_INDEX_NAME']}/docs",
            headers={"api-key": os.environ["SEARCH_KEY"]},
            params={
                "api-version": "2025-08-01-preview",
                "$top": 10,
                "$filter": filter_expr,
                "$select": "key,url,title"
            },
            timeout=30
        )
        index_docs = search_resp.json().get("value", []) if search_resp.status_code == 200 else []
      
        url_map: Dict[str, Dict] = {}
        for doc in index_docs:
            index_url = normalize_url(doc.get("url", ""))
            if index_url:
                url_map[index_url] = doc

        # Deduplicate citations by normalized url
        unique_docs: List[Dict] = []
        doc_key_map: Dict[str, int] = {}

        for doc in citations:
            cit_url = normalize_url(doc.get("url", ""))
      
            match_doc = url_map.get(cit_url)
            if match_doc:
                doc["url"] = match_doc.get("url", "")
                doc["title"] = doc.get("title") or match_doc.get("title", "")
            
            k = cit_url or (doc.get("key") or doc.get("title") or "").lower()
            if k and k not in doc_key_map:
                doc_key_map[k] = len(unique_docs) + 1
                unique_docs.append(doc)

        def replace_citations(text: str, cits: List[Dict]) -> str:
            orig_to_unique = {}
            for idx, d in enumerate(cits):
                k = normalize_url(d.get("url", "")) or (d.get("key") or d.get("title") or "").lower()
                if k in doc_key_map:
                    orig_to_unique[str(idx + 1)] = str(doc_key_map[k])
            return re.sub(r"\[(\d+)\]", lambda m: f"[{orig_to_unique.get(m.group(1), m.group(1))}]", text)

        assistant_reply = replace_citations(assistant_reply, citations)
        valid_nums = {str(i + 1) for i in range(len(unique_docs))}
        assistant_reply = remove_orphan_citations(assistant_reply, valid_nums)
        assistant_reply = _append_reference_links(assistant_reply, unique_docs)

        references: List[Dict] = []
        for i, doc in enumerate(unique_docs):
            idx = i + 1
            resolved_url = (doc.get("url") or "").strip()
            if resolved_url:
                if "?" in resolved_url:
                    if "csf=1" not in resolved_url: resolved_url += "&csf=1"
                    if "web=1" not in resolved_url: resolved_url += "&web=1"
                    if "action=default" not in resolved_url: resolved_url += "&action=default"
                else:
                    resolved_url += "?csf=1&web=1&action=default"
            else:
                resolved_url = "#"

            # NEW: compute a stable, human-friendly label without constructing URLs
            index_doc = url_map.get(normalize_url(doc.get("url", "")))
            title = pick_display_name(doc, index_doc, resolved_url)

            references.append({
                "index": idx,
                "title": title,
                "url": resolved_url
            })
            logging.info("Resolved reference %s → title=%r url=%r", idx, title, resolved_url)

        valid_nums = {str(ref["index"]) for ref in references}
        assistant_reply = remove_orphan_citations(assistant_reply, valid_nums)

        # ---------- Send response ----------
        response_payload = {
            "answer": assistant_reply,
            "history": chat_history + [
                {"role": "user", "content": message},
                {"role": "assistant", "content": assistant_reply}
            ],
            "references": references
        }

        # Log EXACTLY what we send back
        logging.info("Sending response payload: %s", _preview(response_payload))
        print("RESPONSE PAYLOAD:", _preview(response_payload))

        return func.HttpResponse(
            json.dumps(response_payload),
            status_code=200,
            mimetype="application/json"
        )

    except Exception as e:
        logging.exception("Unhandled exception in ask_ai")
        return _error_response("Unhandled exception", 500, extra={"message": str(e), "trace": traceback.format_exc()})
