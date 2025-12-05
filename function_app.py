import logging
import os
import json
import requests
import azure.functions as func
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient
from datetime import datetime
from urllib.parse import quote,urljoin
import base64
import re
import traceback
from typing import List, Dict

logging.basicConfig(level=logging.INFO)
app = func.FunctionApp()

REQUIRED_ENV_VARS = [
    "AI_FOUND_ENDPOINT",
    "AI_FOUND_API_KEY",
    "SEARCH_ENDPOINT",
    "SEARCH_INDEX_NAME",
    "SEARCH_KEY"
]
def get_env_var(key, default=None):
    val = os.environ.get(key, default)
    if val is None:
        raise EnvironmentError(f"Missing required environment variable: {key}")
    return val

@app.function_name(name="TimerTrigger")
@app.schedule(schedule="0 0 * * * *", arg_name="mytimer", run_on_startup=True, use_monitor=False)
def main(mytimer: func.TimerRequest) -> None:
    logging.info("🔁 TimerTrigger sync started.")

    try:
        MAX_FILES = 50
        processed_files = 0
        BLOB_PREFIX = "sharepoint-data"
        PROGRESS_FILE = f"{BLOB_PREFIX}/.progress.json"

        blob_conn = os.environ["BLOB_CONNECTION_STRING"]
        container = os.environ.get("SHAREPOINT_CONTAINER", "rtlgssfinitcontainer2")
        hostname = os.environ["TENANT_NAME"]

        drive_targets = json.loads(json.loads(os.environ["SHAREPOINT_DRIVES"]))

        blob_service = BlobServiceClient.from_connection_string(blob_conn)
        blob_client = blob_service.get_blob_client(container=container, blob=PROGRESS_FILE)

        credential = DefaultAzureCredential()
        token = credential.get_token("https://graph.microsoft.com/.default").token
        headers = {"Authorization": f"Bearer {token}"}

        try:
            progress = json.loads(blob_client.download_blob().readall())
        except:
            progress = {}

        new_progress = {}

        def save_progress():
            blob = blob_service.get_blob_client(container=container, blob=PROGRESS_FILE)
            blob.upload_blob(json.dumps(new_progress, indent=2), overwrite=True)

        def resolve_site_id(site_path):
            resp = requests.get(f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}", headers=headers)
            resp.raise_for_status()
            return resp.json()["id"]

        def resolve_drive_id(site_id, drive_name):
            url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
            resp = requests.get(url, headers=headers)
            resp.raise_for_status()
            for d in resp.json().get("value", []):
                if d["name"].lower() == drive_name.lower():
                    return d["id"]
            raise Exception(f"Drive '{drive_name}' not found")

        def list_all_files(drive_id, folder_id="root", folder_path=""):
            items = []
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children"
            resp = requests.get(url, headers=headers)
            resp.raise_for_status()
            for item in resp.json().get("value", []):
                name = item["name"]
                path = f"{folder_path}/{name}".strip("/")
                if "folder" in item:
                    items += list_all_files(drive_id, item["id"], path)
                else:
                    items.append({
                        "id": item["id"],
                        "name": name,
                        "path": path,
                        "lastModifiedDateTime": item["lastModifiedDateTime"],
                        "webUrl": item.get("webUrl"),
                        "size": item.get("size"),
                        "file": item.get("file", {})
                    })
            return items

        def upload_file(blob_path, content_url):
            nonlocal processed_files
            with requests.get(content_url, headers=headers, stream=True) as r:
                if r.status_code == 200:
                    blob = blob_service.get_blob_client(container=container, blob=blob_path)
                    blob.upload_blob(r.raw, overwrite=True)
                    processed_files += 1
                    logging.info(f"📤 Uploaded: {blob_path}")

        def upload_metadata(blob_path, metadata):
            blob = blob_service.get_blob_client(container=container, blob=blob_path + ".metadata.json")
            blob.upload_blob(json.dumps(metadata), overwrite=True)

        def delete_orphans(blob_paths_expected):
            existing_blobs = blob_service.get_container_client(container).list_blobs(name_starts_with=BLOB_PREFIX)
            for blob in existing_blobs:
                path = blob.name
                if path.endswith(".metadata.json") or path.endswith(".progress.json"):
                    continue
                if path not in blob_paths_expected:
                    blob_service.get_blob_client(container=container, blob=path).delete_blob()
                    logging.info(f"🗑️ Deleted orphan blob: {path}")

        # === Process Each Drive in Order ===
        blob_paths_to_keep = set()
        for entry in drive_targets:
            site_path = entry["site_path"]
            site_name = entry.get("site_name", site_path.strip("/").split("/")[-1])
            drive_name = entry["drive_name"]
            key = f"{site_name}:{drive_name}"

            if processed_files >= MAX_FILES:
                logging.warning(f"⏭️ Skipping drive '{drive_name}' due to global file limit ({MAX_FILES})")
                continue

            logging.info(f"🔄 Processing drive: {drive_name}")
            prev = progress.get(key, {})

            try:
                site_id = resolve_site_id(site_path)
                drive_id = resolve_drive_id(site_id, drive_name)

                files = list_all_files(drive_id)
                new_progress[key] = {}

                for f in files:
                    if processed_files >= MAX_FILES:
                        logging.warning(f"🚧 Max file limit reached mid-drive: {drive_name}")
                        break

                    # Skip large or media files
                    if f["name"].lower().endswith(".mp4") or f["size"] > 100 * 1024 * 1024:
                        logging.warning(f"⏩ Skipping large/media file: {f['name']} ({f['size']} bytes)")
                        continue

                    blob_path = f"{BLOB_PREFIX}/{site_name}/{drive_name}/{f['path']}"
                    blob_paths_to_keep.add(blob_path)

                    last_mod = f["lastModifiedDateTime"]
                    if f["name"] not in prev or prev[f["name"]] != last_mod:
                        upload_file(blob_path, f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{f['id']}/content")
                        upload_metadata(blob_path, f)

                    new_progress[key][f["name"]] = last_mod

                if processed_files < MAX_FILES:
                    logging.info(f"✅ Finished drive: {drive_name}")

            except Exception as e:
                logging.warning(f"⚠️ Skipping drive '{drive_name}': {e}")

        delete_orphans(blob_paths_to_keep)
        save_progress()
        logging.info("✅ Sync complete.")

    except Exception:
        logging.exception("❌ TimerTrigger execution failed")



def load_system_prompt(path="system_prompt.md"):
    try:
        if not os.path.isabs(path):
            path = os.path.join(os.path.dirname(__file__), path)
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        logging.error(f"System prompt file not found at {path}")
        return "System prompt file not found."
    

def decode_blob_url(chunk_or_parent_id: str) -> str:
    """
    Decode blob URL from chunk_id or parent_id and convert to SharePoint path
    """
    try:
        # Extract the encoded part (usually after last underscore)
        encoded = chunk_or_parent_id.split("_")[-1] if "_" in chunk_or_parent_id else chunk_or_parent_id
        
        # Add padding for base64 decoding
        padded = encoded + '=' * ((4 - len(encoded) % 4) % 4)
        
        # Decode base64
        decoded = base64.b64decode(padded).decode("utf-8")
        
        # Remove blob storage prefix if present
        if decoded.startswith("sharepoint-data/"):
            decoded = decoded.replace("sharepoint-data/", "")
        
        # Remove site name prefix if present
        path_parts = decoded.split("/")
        if len(path_parts) > 2:
            # Skip site name and drive name, keep the actual file path
            decoded = "/".join(path_parts[2:])
        
        # URL encode the path
        return quote(decoded, safe=':/')
    except Exception as e:
        logging.warning(f"Failed to decode blob URL from {chunk_or_parent_id}: {e}")
        return None

def _generate_sharepoint_url(doc: dict) -> str:
    try:
        base_url = os.environ.get("SHAREPOINT_BASE_URL", "").rstrip("/")
        if not base_url.startswith("http"):
            logging.warning("❌ SHAREPOINT_BASE_URL is not set or invalid")
            return None

        logging.info(f"🔗 Generating SharePoint URL for document: {doc.get('title')}")

        # 1. Direct URL from metadata (must start with http)
        doc_url = doc.get("url", "")
        if doc_url and doc_url.startswith("http"):
            logging.info("📎 Using direct full URL from document metadata")
            # ensure ?csf=1&web=1 is appended
            if "?" in doc_url:
                if "csf=1" not in doc_url:
                    doc_url += "&csf=1"
                if "web=1" not in doc_url:
                    doc_url += "&web=1"
                if "action=default" not in doc_url:
                    doc_url += "&action=default"
            else:
                doc_url += "?csf=1&web=1&action=default"
            return doc_url

        # 2. Try decoding from chunk_id or parent_id
        chunk_id = doc.get("chunk_id") or doc.get("parent_id")
        if chunk_id:
            logging.info(f"🧬 Decoding chunk_id: {chunk_id}")
            decoded_path = decode_blob_url(chunk_id)
            logging.info(f"🛣️ Decoded SharePoint path: {decoded_path}")

            if decoded_path:
                encoded_path = quote(decoded_path, safe=':/')
                full_url = urljoin(base_url + "/", encoded_path)
                # Append ?web=1
                if "?web=1" not in full_url:
                    full_url += "?web=1"
                return full_url


        # 3. Fallback: Use heuristics based on title keywords
        title = doc.get("title", "")
        title_lower = title.lower()
        directory_map = {
            "compliance": "Compliance",
            "trans4m": "TRANS4M Training Materials",
            "gtrs": "GTRS Calendar Monthly Newsletters",
            "group finance": "Group Finance Referenced Applications",
            "referenced applications": "Group Finance Referenced Applications"
        }

        for keyword, folder in directory_map.items():
            if keyword in title_lower:
                full_path = f"{folder}/{title}"
                encoded_path = quote(full_path, safe="/")
                constructed_url = urljoin(base_url + "/", encoded_path)
                if "?web=1" not in constructed_url:
                    constructed_url += "?web=1"
                return constructed_url

        # 4. Final fallback: Group Finance Referenced Applications
        full_path = f"Group Finance Referenced Applications/{title}"
        encoded_path = quote(full_path, safe="/")
        
        default_url = urljoin(base_url + "/", encoded_path)
        if "?web=1" not in default_url:
            default_url += "?web=1"
        logging.info(f"🧾 Final fallback URL: {default_url}")
        return default_url


    except Exception as e:
        logging.warning(f"❌ Failed to generate SharePoint URL: {e}")
        return None
    
def remove_orphan_citations(text: str, valid_numbers: set) -> str:
    return re.sub(r'\[(\d+)\]', lambda m: f"[{m.group(1)}]" if m.group(1) in valid_numbers else "", text)

@app.function_name(name="HttpAskAI")
@app.route(route="ask-ai", auth_level=func.AuthLevel.FUNCTION)
def ask_ai(req: func.HttpRequest) -> func.HttpResponse:
    try:
        req_body = req.get_json()
        message = req_body.get("message")
        chat_history = req_body.get("history", [])

        if not message:
            return _error_response("Missing 'message' in request body", 400)

        missing = [key for key in REQUIRED_ENV_VARS if not os.getenv(key)]
        if missing:
            return _error_response(f"Missing environment variable(s): {missing}", 500)
        

   

        system_message = {
            "role": "system",
            "content": load_system_prompt()
        }

        # system_message = {
        #     "role": "system",
        #     "content": (
        #         "You are a helpful AI assistant with access to RTL Group Finance IT documents. "
        #         "Provide accurate information about the content in the selected files and reply in a formal tone. "
        #         "These documents were stored in SharePoint, uploaded to Azure Blob Storage, and indexed in Azure Cognitive Search. "
        #         "Each result includes metadata like title and a url decoded from chunk_id or parent_id.\n\n"
        #         "Always cite documents using [1], [2], etc. Never use [doc1], [doc2], or similar non-numeric formats. Do not output a list of document titles at the top of your reply."
        #         "Topics include:\n"
        #         "- Group Finance Referenced Applications\n"
        #         "- Compliance\n"
        #         "- TRANS4M Training Materials\n"
        #         "- GTRS Calendar / Newsletters\n\n"
        #         "Guidelines:\n"
        #         "- Use only the retrieved documents as your source of truth\n"
        #         "- When citing documents, use the format [1], [2], etc. and the citation in form superscript (these will be converted to simplified references)\n"
        #         "- Cite after relevant sentences or paragraphs\n"
        #         "- When answering in german, use 'du' form instead of 'sie' \n"
        #         "- Use clear, concise language and numbered bullet points when listing items\n"
        #         "- Do not fabricate citations\n\n"
        #         "✅ Example:\n"
        #         "The Automatic Payment Program processes outgoing payments [1]. Key points include vendor selection and payment parameters [2]."
        #     )
        # }

        messages = [system_message] + chat_history + [{"role": "user", "content": message}]

        ai_response = requests.post(
            os.environ["AI_FOUND_ENDPOINT"],
            headers={
                "api-key": os.environ["AI_FOUND_API_KEY"],
                "Content-Type": "application/json"
            },
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
                            "semantic_configuration": "rag-rtlgss-finit-prod-semantic-configuration",
                            "query_type": "semantic",
                            "fields_mapping": {},
                            "in_scope": True,
                            "filter": None,
                            "strictness": 3,
                            "top_n_documents": 5,
                            "authentication": {
                                "type": "api_key",
                                "key": os.environ["SEARCH_KEY"]
                            }
                        }
                    }
                ]
            }
        )

        if ai_response.status_code != 200:
            return _error_response("AI service request failed", ai_response.status_code, extra={"message": ai_response.text})

        result = ai_response.json()

        assistant_reply = result["choices"][0]["message"]["content"]
        citations = result["choices"][0]["message"].get("context", {}).get("citations", [])

        # Step 1: Deduplicate citations by document (prefer URL, fallback to title)
        unique_docs = []
        doc_key_map = {}  # Maps doc key to its assigned citation number

        def doc_key(doc):
            return (doc.get("url") or "").lower() or doc.get("title", "").lower()

        for doc in citations:
            key = doc_key(doc)
            if key and key not in doc_key_map:
                doc_key_map[key] = len(unique_docs) + 1  # Citation number starts at 1
                unique_docs.append(doc)

        # Step 2: Replace all citations in the answer text with their unique number
        def replace_citations(text, citations):
            orig_to_unique = {}
            for idx, doc in enumerate(citations):
                key = doc_key(doc)
                if key in doc_key_map:
                    orig_to_unique[str(idx + 1)] = str(doc_key_map[key])
            def repl(m):
                orig = m.group(1)
                return f"[{orig_to_unique.get(orig, orig)}]"
            return re.sub(r'\[(\d+)\]', repl, text)

        assistant_reply = replace_citations(assistant_reply, citations)

        # Step 3: Remove orphan citations (in case any remain)
        valid_citation_numbers = {str(i + 1) for i in range(len(unique_docs))}
        assistant_reply = remove_orphan_citations(assistant_reply, valid_citation_numbers)

        # Step 4: Append only unique citations with links
        assistant_reply = _append_reference_links(assistant_reply, unique_docs)

        # Step 5: Build references array
        references = []
        for i, doc in enumerate(unique_docs):
            index = i + 1
            title = doc.get("title")
            if not title:
                url = doc.get("url") or _generate_sharepoint_url(doc)
                title = url.split("/")[-1].split("?")[0] if url else f"Document {index}"
            references.append({
                "index": index,
                "title": title,
                "url": _generate_sharepoint_url(doc)
            })


        # 🔒 Recompute valid references after final filtering
        valid_citation_numbers = {str(ref["index"]) for ref in references}

        # 🧹 Remove orphan citations again (in case answer included unused numbers)
        assistant_reply = remove_orphan_citations(assistant_reply, valid_citation_numbers)


        return func.HttpResponse(
            json.dumps({
                "answer": assistant_reply,
                "history": chat_history + [
                    {"role": "user", "content": message},
                    {"role": "assistant", "content": assistant_reply}
                ],
                "references": references   # <-- Add this line
            }),
            status_code=200,
            mimetype="application/json"
        )

    except Exception as e:
        logging.exception("❌ Unhandled exception in ask_ai")
        return _error_response("Unhandled exception", 500, extra={
            "message": str(e),
            "trace": traceback.format_exc()
        })

def _append_reference_links(answer: str, docs: List[Dict]) -> str:
    if not docs or not isinstance(docs, list):
        return answer

    # Fix joined citations like [1][2] and [1][2][3]
    answer = re.sub(r'\[(\d+)\]\[(\d+)\]', r'[\1], [\2]', answer)
    answer = re.sub(r'(\[\d+\])(?=\[\d+\])', r'\1, ', answer)

    # Extract used reference numbers
    used_refs = set(re.findall(r'\[(\d+)\]', answer))
    if not used_refs:
        return answer

    # Prepare references section
    references_section = "\n\n**Sources:**\n"
    base_url = os.environ.get("SHAREPOINT_BASE_URL", "").rstrip("/")

    for ref in sorted(used_refs, key=int):
        i = int(ref) - 1
        if i < len(docs):
            raw_url = docs[i].get("url") or _generate_sharepoint_url(docs[i]) or ""


            if raw_url:
                # If URL is relative, make it absolute
                if not raw_url.startswith("http"):
                    # Safely encode path for SharePoint
                    safe_path = quote(raw_url.lstrip("/"), safe=':/?=&')
                    full_url = urljoin(base_url + "/", safe_path)
                else:
                    full_url = raw_url

                references_section += f"[{ref}]: {full_url}\n"
                logging.debug(f"🔗 Citation [{ref}] → {full_url}")
            else:
                references_section += f"[{ref}]: #\n"
                logging.warning(f"⚠️ Missing URL for citation [{ref}]")
        else:
            references_section += f"[{ref}]: #\n"
            logging.warning(f"⚠️ Reference index [{ref}] out of bounds")

    # return answer.strip() + references_section
    return answer.strip() + "\n\n" + references_section.strip() + "\n"



def _error_response(message: str, status_code: int, extra: dict = None) -> func.HttpResponse:
    payload = {"error": message}
    if extra:
        payload.update(extra)
    return func.HttpResponse(
        json.dumps(payload),
        status_code=status_code,
        mimetype="application/json"
    )
