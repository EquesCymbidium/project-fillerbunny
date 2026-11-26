# Fillerbunny

Turn a short animation prompt into a downloadable PowerPoint deck and MP4 preview.

- GitHub Pages publishes the interface from the `gh-pages` branch.
- Render runs the Python generator from [`render.yaml`](./render.yaml).
- The cloud generator reuses the prompt parser and writes native PowerPoint animation timelines directly in Open XML; it does not need PowerPoint COM or Ollama at runtime. Fox and dog artwork is included, while other entities use labeled shapes.

See [Web/README.md](./Web/README.md) for local development and deployment steps.
