# watsonx Dashboard — Backend Server

Node.js proxy that sits between the React frontend and the IBM watsonx.ai API.  
It keeps the IBM API key off the browser entirely.

---

## Folder structure

```
server/
  index.js          ← Express app + /api/chat endpoint
  iamToken.js       ← IBM Cloud IAM token fetcher (with in-memory cache)
  watsonx.js        ← Prompt builder + watsonx.ai REST caller
  package.json
  .env.example      ← Copy this to .env and fill in your keys
```

---

## Setup

### 1. Install dependencies
```bash
cd server
npm install
```

### 2. Create your .env file
```bash
cp .env.example .env
```
Open `.env` and fill in your real values:

| Variable | Description |
|---|---|
| `IBM_CLOUD_API_KEY` | IBM Cloud IAM API key |
| `WATSONX_PROJECT_ID` | watsonx.ai project ID |
| `WATSONX_MODEL_ID` | Model to use (default: `ibm/granite-13b-instruct-v2`) |
| `WATSONX_REGION` | Region (default: `us-south`) |
| `PORT` | Server port (default: `3001`) |

### 3. Start the server

**Development** (auto-restarts on file changes — Node 18+ built-in watcher):
```bash
npm run dev
```

**Production**:
```bash
npm start
```

The server listens on `http://127.0.0.1:3001` (localhost only).

---

## API

### `POST /api/chat`

Request body:
```json
{
  "question": "How many P1 tickets are breached?",
  "tickets": [ ...filtered incident objects from the frontend... ]
}
```

Success response:
```json
{
  "answer": "There are 3 breached P1 tickets: INC001, INC002, INC003."
}
```

Error response:
```json
{
  "error": "The AI service encountered an error. Please try again shortly."
}
```

### `GET /health`
Returns `{ "status": "ok" }` — useful to confirm the server is running.

---

## Security

- Binds to `127.0.0.1` only — never exposed on `0.0.0.0`.
- CORS restricted to `http://localhost:5173` (Vite dev server).
- API key lives in `.env` only — never in code or logs.
- Request body capped at 2 MB.
- Inputs validated (type, length, non-empty) before calling the AI.
- Stack traces are logged server-side only; clients receive generic error messages.

---

## Running both frontend and backend together

Open two terminals:

```bash
# Terminal 1 — frontend
cd WatsonX
npm run dev

# Terminal 2 — backend
cd WatsonX/server
npm run dev
```
