# Upload This Project to GitHub

## Option 1: GitHub Website

1. Sign in to GitHub.
2. Click the `+` icon.
3. Select **New repository**.
4. Repository name:

```text
excel-essential-guide
```

5. Description:

```text
A practical Excel formula, table, chart, and reporting guide with sample workbooks.
```

6. Choose **Public**.
7. Do not add another README, license, or `.gitignore`; they are already included.
8. Click **Create repository**.
9. Extract the ZIP package.
10. On the repository page, choose **uploading an existing file**.
11. Drag all files and folders into GitHub.
12. Commit with:

```text
Initial Excel essential guide
```

## Option 2: PowerShell and Git

Extract the ZIP, then open PowerShell inside the repository folder:

```powershell
git init
git add .
git commit -m "Initial Excel essential guide"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/excel-essential-guide.git
git push -u origin main
```

Replace `YOUR-USERNAME` with your GitHub username.
