# Philadelphia PA Address Search Tool - GitHub Scheduled Version

This folder is ready to upload to GitHub.

## Files

- `Philadelphia_PA_AddressSearch_Tool.py` - main Selenium scraper
- `PA.txt` - input file; keep one input/address per line
- `requirements.txt` - Python packages
- `.github/workflows/pa_scheduled_run.yml` - GitHub Actions scheduled workflow

## How to edit input

Open `PA.txt` and add one record per line:

```text
123 MAIN ST PHILADELPHIA PA
456 MARKET ST PHILADELPHIA PA
```

## Default schedule

The workflow is set to run daily at **10:30 AM India time**:

```yaml
- cron: "30 10 * * *"
  timezone: "Asia/Kolkata"
```

Change the cron value if you want a different time.

## Push commands

Open PowerShell inside this folder and run:

```powershell
git init
git add .
git commit -m "Add Philadelphia PA scheduled tool"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git
git push -u origin main
```

## Manual test in GitHub

1. Open your GitHub repository.
2. Go to **Actions**.
3. Open **Philadelphia PA Scheduled Run**.
4. Click **Run workflow**.
5. After it completes, download the artifact named **Philadelphia-PA-Output**.

## Output

The Excel file will be uploaded as a GitHub Actions artifact. It will not automatically appear inside your repository files.

## Note

If the Philadelphia website blocks GitHub cloud traffic, use a GitHub self-hosted runner on your own Windows machine.
