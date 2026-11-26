# Fillerbunny web generator

The site accepts an animation prompt and provides a `.pptx` scene deck plus an `.mp4` preview. It reuses the project's rule-based prompt parser. The public generator is cloud-compatible: PowerPoint COM and local Ollama are not used. Entities are rendered as labeled shapes rather than the original icon artwork. The PowerPoint file contains a slide snapshot after each action; the MP4 visualizes the actions. Native PowerPoint animation fidelity is not preserved. A short built-in entity list keeps common prompts working without publishing the project's icon artwork; the full local icon directory still expands entity recognition on the original project.

## Run locally

From the project root:

```powershell
python -m pip install -r Web\requirements.txt
python -m uvicorn Web.backend.app:app --reload
```

Open `http://127.0.0.1:8000`. The app needs `Product\Blackbox\dispatcher.py`; `Library\Icons` is optional and is not included in the public deployment.

## Public deployment (free tiers)

1. Put the **project root contents** in a GitHub repository. The website workflow assumes `Web` and `.github` are at the repository root.
2. In GitHub, enable **Settings → Pages → Build and deployment → GitHub Actions**. The workflow publishes `Web` as a static site.
3. Create a Render web service from the same GitHub repository using the root `render.yaml` Blueprint. Select the free instance plan. The service sleeps when idle on free tiers; its first request after sleeping can take a minute. Generated downloads are temporary and expire after one hour.
4. Open the Render service URL, confirm `/health` returns `{"status":"ok"}`, then open the GitHub Pages site. Under **Generator connection**, enter the Render service URL and save it. The URL is kept only in that browser.
5. To restrict browser access, set the Render environment variable `CORS_ORIGINS` to the exact GitHub Pages origin (for example `https://OWNER.github.io`). CORS is not authentication; this generator is public and rate-limited to five generation requests per client IP per minute.

The free hosting provider and GitHub Pages are separate: Pages hosts the interface; Render runs Python generation. Do not put credentials in the front end.
