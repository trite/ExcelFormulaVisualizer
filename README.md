# Excel Formula Visualizer

This tool is in very early development right now. **It is very important that you do NOT save over your original document if downloading a copy with metadata attached.**

🚀 **Live Demo**: [https://trite.github.io/ExcelFormulaVisualizer/](https://trite.github.io/ExcelFormulaVisualizer/)

## Deployment

This project is automatically deployed to GitHub Pages when changes are pushed to the `main` branch.

### Automatic Deployment

The GitHub Actions workflow (`.github/workflows/deploy.yml`) handles deployment automatically:

- Triggers on every push to `main`
- Can also be triggered manually from the Actions tab
- Builds the project and deploys to GitHub Pages

### Manual Deployment

To deploy manually using the gh-pages package:

1. Install gh-pages: `npm install -D gh-pages`
2. Run: `npm run deploy`

### First-Time Setup

To enable GitHub Pages for this repository:

1. Go to repository Settings → Pages
2. Under "Build and deployment", set Source to "GitHub Actions"
3. Push to `main` branch to trigger the first deployment
