## 1. [Download the Password Sync Support Tool](https://github.com/google/password-sync-support-tool/releases/download/2.0.3/PasswordSyncSupportTool.vbs)

### 2. Upload the file it creates to the [log analyzer](https://toolbox.googleapps.com/apps/loganalyzer/?productid=gaps)

Learn more about [troubleshooting Password Sync](https://support.google.com/a/answer/2622457), and about [Password Sync logs and error codes](https://support.google.com/a/answer/3296820).

---

This tool collects logs and information from all Domain Controllers running [Password Sync](https://support.google.com/a/answer/2611859) in order to allow reviewing them all in a single place to make troubleshooting easier. It will create a ZIP file on your Desktop when it's finished.

Notes:
* If you have multiple domains in your forest, you need to run the support tool while logged in as a user from the domain you want to investigate. It fetches the logs from all DCs in your domain, not across the entire forest.
  * You don't have to run it from a DC, you can run it from any domain member computer (as long as you're logged in as a Domain Admin), but it's better to run it from a DC that's affected by the issue you want to investigate.
  * Make sure that you have unblocked network connectivity between all writable DCs in your domain. If you don't, some data would be missing.
* If you can't start the support tool:
  1. Right click the file you downloaded (`PasswordSyncSupportTool.vbs`).
  2. Click "Properties".
  3. Click the "Unblock" checkbox at the bottom.
  4. Click "OK".
  5. Try running the support tool again
* If you have a lot of DCs in your domain, it could take a long time for the support tool to run. That's ok, just let it run until it finishes.
* Make sure you don't click the window while it's running. If you do, it could go into "Select" mode, which pauses the run. Make sure that the title of the window doesn't have "Select" in it. If it does, just press the Escape key on your keyboard.

It's built using VBScript for compatibility with all Windows versions.

> Password Sync was previously known as "Google Apps Password Sync" and "G Suite Password Sync". This support tool was previously known as GAPSTool and GSPSTool. 
