# Contact Refresh

A web tool that compares old vs. new contact data (title/company) and flags what changed, using Claude for the comparison. Colors are inspired by hottopics.ht (navy header, pink/yellow accents) — swap the hex values at the top of `src/App.jsx` if you have exact brand colors.

## Deploy it (free, ~10 minutes, no coding required)

**1. Get an Anthropic API key**
Go to https://console.anthropic.com, sign in, create an API key. Keep this private — you'll paste it into Vercel in step 4, never into the code itself.

**2. Put this project on GitHub**
- Create a free GitHub account if you don't have one: https://github.com
- Create a new repository (e.g. `contact-refresh`)
- Upload all the files in this folder to that repository (GitHub's web uploader works fine — drag and drop the files in)

**3. Deploy to Vercel**
- Go to https://vercel.com and sign up (free) using your GitHub account
- Click "Add New Project", select the `contact-refresh` repository you just created
- Leave the default settings (Vercel auto-detects this as a Vite project) and click Deploy

**4. Add your API key**
- Once deployed, go to your project's Settings → Environment Variables in Vercel
- Add a variable named `ANTHROPIC_API_KEY` with your key from step 1 as the value
- Go to the Deployments tab and click "Redeploy" so the key takes effect

**5. Done**
Vercel gives you a live URL like `contact-refresh-yourname.vercel.app`. That's your website — share it, bookmark it, or later attach your own domain under Settings → Domains.

## Running it locally first (optional)
If you want to test before deploying:
```
npm install
npm run dev
```
This runs the interface, but the `/api/compare` endpoint only works once deployed to Vercel (or via `vercel dev` if you install the Vercel CLI).

## Files
- `src/App.jsx` — the interface and all the logic
- `api/compare.js` — the serverless function that talks to Claude (keeps your API key private)
- `index.html`, `src/main.jsx`, `vite.config.js` — standard project scaffolding, no need to touch these
