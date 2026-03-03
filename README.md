# PixelBug Calendar

Simple browser-based scheduling prototype for PixelBug's lighting design team.

## Run

Open [/Users/Rahul/Documents/Playground/index.html](/Users/Rahul/Documents/Playground/index.html) in a browser.

## Local Preview Server

Run:

```bash
npm run dev
```

Then open [http://127.0.0.1:4173](http://127.0.0.1:4173).

## Cloud Deploy

This project is prepared for Render deployment using the existing SQLite backend with a persistent disk.

Deployment files:

- [render.yaml](/Users/Rahul/Documents/Playground/render.yaml)
- [.env.example](/Users/Rahul/Documents/Playground/.env.example)
- [.gitignore](/Users/Rahul/Documents/Playground/.gitignore)

### Render Steps

1. Push this project to GitHub.
2. In Render, create a new Blueprint and connect the repository.
3. Render will detect [render.yaml](/Users/Rahul/Documents/Playground/render.yaml).
4. Set these environment variables in Render:
   - `PIXELBUG_BASE_URL`
   - `RESEND_API_KEY`
   - `PIXELBUG_EMAIL_FROM`
5. Deploy the service.

### Deployment Note

This setup keeps SQLite on a Render persistent disk. It is suitable for a single running web service. The next infrastructure upgrade after this is migrating the backend to managed PostgreSQL if you want a more scalable multi-instance deployment.

## Included

- Admin, crew, and view-only login types
- Admin approval workflow for crew and view-only accounts
- Shared calendar with crew color coding
- Multiple crew assignments per show
- Show amount visible only to admins
- Operator amount visible to admins and the assigned crew member only
- SQLite-backed shared data store
- Password change and password reset flows
- Admin approval flow for new accounts

## Note

The browser now stores only UI state locally. Shared users and shows are stored in `data/pixelbug.db`.

On first launch, create the initial admin account from the login panel, then add crew/view-only users through the app workflow.

## Email Delivery

To send real password reset emails, set:

```bash
export RESEND_API_KEY=your_key_here
export PIXELBUG_EMAIL_FROM="PixelBug <no-reply@yourdomain.com>"
export PIXELBUG_BASE_URL="http://127.0.0.1:4173"
```

Without those env vars, the app writes outgoing emails to:

```bash
/Users/Rahul/Documents/Playground/data/email-outbox.log
```

That fallback is useful for local testing.
# pixelbug
