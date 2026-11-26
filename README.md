# Fillerbunny

Turn a short animation prompt into a downloadable PowerPoint deck and MP4 preview.

- GitHub Pages publishes the interface from the `gh-pages` branch.
- Render runs the Python generator from [`render.yaml`](./render.yaml).
- The cloud generator reuses the prompt parser and renders labeled shapes; it does not depend on PowerPoint, Ollama, or the local icon artwork.

See [Web/README.md](./Web/README.md) for local development and deployment steps.
