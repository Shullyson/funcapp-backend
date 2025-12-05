# Finance IT - Azure Functions App

A serverless Azure Functions application that synchronizes SharePoint documents to Azure Blob Storage and provides AI-powered document querying capabilities for RTL Group Finance IT operations.

## 🚀 Features

### 📅 Automated Document Synchronization (Timer Trigger)
- **Schedule**: Runs every an hour (`0 *** * * * *`)
- **Source**: SharePoint drives from RTL Group Finance IT site
- **Destination**: Azure Blob Storage with organized folder structure
- **Smart Sync**: Only updates files that have been modified since last sync
- **Progress Tracking**: Maintains sync state to avoid unnecessary re-processing
- **File Filtering**: Automatically skips large files (>100MB) and media files (.mp4)
- **Orphan Cleanup**: Removes files from blob storage that no longer exist in SharePoint

### 🤖 AI-Powered Document Query (HTTP Trigger)
- **Endpoint**: `/ask-ai`
- **Functionality**: Natural language querying of synchronized documents
- **AI Integration**: Uses Azure OpenAI (GPT-4.1-mini) with Azure Cognitive Search
- **Smart Citations**: Automatically generates reference links to source documents
- **Chat History**: Maintains conversation context for follow-up questions

## 🛠️ Technology Stack

- **Runtime**: Python 3.10
- **Platform**: Azure Functions (v4)
- **Authentication**: Azure DefaultAzureCredential (Managed Identity)
- **Storage**: Azure Blob Storage
- **Search**: Azure Cognitive Search with semantic configuration
- **AI**: Azure OpenAI (GPT-4.1-mini)
- **APIs**: Microsoft Graph API for SharePoint integration

## 📋 Prerequisites

- Azure subscription with the following services:
  - Azure Functions
  - Azure Blob Storage
  - Azure Cognitive Search
  - Azure OpenAI
- SharePoint Online access to RTL Group tenant
- Python 3.10 or later
- Azure Functions Core Tools

## ⚙️ Environment Variables

The following environment variables must be configured:

### Required for both functions:
```
FUNCTIONS_WORKER_RUNTIME=python
AzureWebJobsStorage=<Azure Storage connection string>
```

### For SharePoint Synchronization (Timer Trigger):
```
BLOB_CONNECTION_STRING=<Azure Blob Storage connection string>
SHAREPOINT_CONTAINER=<Blob container name>
TENANT_NAME=<SharePoint tenant name>
SHAREPOINT_DRIVES=<JSON array of drive configurations>
```

### For AI Query Function (HTTP Trigger):
```
AI_FOUND_ENDPOINT=<Azure OpenAI endpoint URL>
AI_FOUND_API_KEY=<Azure OpenAI API key>
SEARCH_ENDPOINT=<Azure Cognitive Search endpoint>
SEARCH_INDEX_NAME=<Search index name>
SEARCH_KEY=<Azure Cognitive Search API key>
```

## 🏗️ Project Structure

```
FinIT-Agent-Backend/
├── function_app.py          # Main application with both functions
├── requirements.txt         # Python dependencies
├── host.json               # Functions host configuration
├── local.settings.json     # Local development settings
└── __pycache__/           # Python bytecode cache
```

## 🚀 Local Development

1. **Clone the repository**
   ```bash
   git clone https://RTLGroup-SharePoint@dev.azure.com/rtlgroup-sharepoint/DWS%20-%20Finance%20IT%20Agent/_git/FinIT-Agent-Backend
   cd FinIT-Agent-Backend
   ```

2. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure local settings**
   - Copy `local.settings.json.example` to `local.settings.json`
   - Update all environment variables with your Azure resource details

4. **Run locally**
   ```bash
   func start
   ```

## 📨 API Usage

### Query Documents with AI

**Endpoint**: `POST /api/ask-ai`

**Request Body**:
```json
{
  "message": "What are the compliance requirements for financial reporting?",
  "history": [
    {
      "role": "user", 
      "content": "Previous question"
    },
    {
      "role": "assistant", 
      "content": "Previous response"
    }
  ]
}
```

**Response**:
```json
{
  "answer": "Based on the compliance documents...",
  "history": [
    // Updated conversation history
  ]
}
```



### Azure Cognitive Search Configuration
- **Semantic Configuration**: `your-semantic-configuration`
- **Query Type**: Semantic search
- **Top Documents**: 5 results per query
- **Strictness**: Level 3

## 🔍 Monitoring and Logging

The application uses structured logging with emoji indicators:

- 🔁 Sync operations
- 📤 File uploads  
- ✅ Successful completions
- ⚠️ Warnings and skipped items
- ❌ Errors and failures
- 🗑️ Cleanup operations

## 🚀 Deployment

### Using Azure Functions Core Tools
```bash
func azure functionapp publish <function-app-name>
```

## 🛡️ Security Considerations

- Uses Azure Managed Identity for authentication
- API keys stored as application settings (encrypted at rest)
- Function-level authentication required for HTTP endpoints
- SharePoint access limited to configured drives only
- File size limits prevent abuse (100MB max)

### Troubleshooting
- Check Azure Function logs for detailed error information
- Verify SharePoint permissions if sync fails
- Ensure all environment variables are properly configured
- Monitor Azure service health for dependencies

### Key APIs Used
- **Microsoft Graph API v1.0** - SharePoint document access
- **Azure Blob Storage SDK** - File storage operations
- **Azure Cognitive Search** - Document indexing and semantic search
- **Azure OpenAI** - GPT-4.1-mini chat completions with data sources

## � References

### Documentation
- [Azure Functions Python Developer Guide](https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference-python)
- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Azure Blob Storage Python SDK](https://docs.microsoft.com/en-us/azure/storage/blobs/storage-quickstart-blobs-python)
- [Azure Cognitive Search REST API](https://docs.microsoft.com/en-us/rest/api/searchservice/)
- [Azure OpenAI Service Documentation](https://docs.microsoft.com/en-us/azure/cognitive-services/openai/)
- [DefaultAzureCredential](https://docs.microsoft.com/en-us/python/api/azure-identity/azure.identity.defaultazurecredential)
- [Managed Identity for Azure Functions](https://docs.microsoft.com/en-us/azure/app-service/overview-managed-identity)
- [Azure Functions Security](https://docs.microsoft.com/en-us/azure/azure-functions/security-concepts)
