# Fillerbunny web generator

The site accepts an animation prompt and provides a native-animated `.pptx` plus an `.mp4` preview. PowerPoint animations start automatically in order. Both downloaded files have a shared hexadecimal timestamp suffix to distinguish separate generations. It reuses the project's rule-based prompt parser and shares its 2,000-character and 10-action limits with the local CLI. The local CLI uses the same deterministic parser by default; its optional `--llm` mode can produce different action plans for unsupported prompts. The public generator writes PowerPoint animation timelines directly in Open XML, so it does not need PowerPoint COM or local Ollama at runtime. White slides and video frames use the project's fox and dog silhouettes; other entities are rendered as labeled shapes. A short built-in entity list keeps common prompts working without publishing the full icon library.

## Run locally

From the project root:

```powershell
python -m pip install -r Web\requirements.txt
python -m uvicorn Web.backend.app:app --reload
```

Open `http://127.0.0.1:8000`. The app needs `Product\Blackbox\dispatcher.py`; `Library\Icons` is optional and is not included in the public deployment.

## Public deployment (free tiers)

1. In GitHub, set **Settings → Pages → Build and deployment → Deploy from a branch**, select `gh-pages` and `/ (root)`.
2. Create a Render web service from the repository using the root `render.yaml` Blueprint. Select the free instance plan. The service sleeps after 15 minutes idle; its first request after sleeping can take about a minute. Generated downloads are kept for up to one hour, but Render's free ephemeral filesystem can clear them sooner when the service sleeps or restarts.
3. Open the Render service URL, confirm `/health` returns `{"status":"ok"}`, then open the GitHub Pages site. Under **Generator connection**, enter the Render service URL and save it. The URL is kept only in that browser.
4. To update the Pages site after changing `Web`, run `git subtree push --prefix Web origin gh-pages`.
5. To restrict browser access, set the Render environment variable `CORS_ORIGINS` to the exact GitHub Pages origin (for example `https://OWNER.github.io`). CORS is not authentication; this generator is public and rate-limited to five generation requests per client IP per minute.

The free hosting provider and GitHub Pages are separate: Pages hosts the interface; Render runs Python generation. Do not put credentials in the front end.
