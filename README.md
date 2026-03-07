# bridge_analyzer
Analyze club bridge session results files (either pbn or BWS), update my google sheet automatically.
Used for session post mortems with my partner:

<img width="2030" height="1074" alt="image" src="https://github.com/user-attachments/assets/81d3dc41-6b17-4693-bcbc-62258e6d01fd" />

## Renewing the Google API credentials

The app authenticates via a **service account** JSON key file stored at:

```
~/.local/share/google-sheets-service-account.json
```

Service account keys don't expire on their own, but if the key is revoked or lost you'll need to generate a new one:

1. Go to the [Google Cloud Console](https://console.cloud.google.com/) and select the project.
2. Navigate to **IAM & Admin > Service Accounts**.
3. Find the service account used by this app and click on it.
4. Go to the **Keys** tab and click **Add Key > Create new key**.
5. Choose **JSON** and click **Create** — this downloads a new key file.
6. Replace the existing credentials file with the downloaded file:
   ```sh
   mv ~/Downloads/<downloaded-key>.json ~/.local/share/google-sheets-service-account.json
   ```
7. Make sure the service account still has **Editor** access to the target Google Sheet (share the sheet with the service account's email address if needed).

