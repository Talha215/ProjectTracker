# Project Dashboard

A tile-based project management dashboard built with React + Vite. Data persists in localStorage.

## Setup

```bash
npm install
npm run dev
```

Opens at `http://localhost:5173/project-dashboard/`

## Deploy to GitHub Pages

1. Create a repo on GitHub called `project-dashboard`
2. Push this code to it:
   ```bash
   git init
   git add .
   git commit -m "initial commit"
   git remote add origin git@github.com:YOUR_USERNAME/project-dashboard.git
   git push -u origin main
   ```
3. Deploy:
   ```bash
   npm run deploy
   ```
4. In your GitHub repo settings, go to **Pages** and confirm the source is set to the `gh-pages` branch.
5. Your dashboard will be live at `https://YOUR_USERNAME.github.io/project-dashboard/`

## Updating

After making changes, just run `npm run deploy` again to push a new build.

## Using with Claude Code

Open this directory in your terminal and run `claude` to iterate on the dashboard with Claude Code.
