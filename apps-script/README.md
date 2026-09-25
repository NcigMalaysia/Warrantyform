# Warranty form Apps Script update

Google Apps Script can save a POST while the browser cannot read its redirected
response. The form now sends the POST once and checks the saved Claim ID with
the read-only `doGet` endpoint before showing a success message.

1. In the existing warranty Apps Script project, add `doGet` from
   [Code.gs](Code.gs) above the existing `doPost`. If the previous version was
   never installed, replace `doPost` and `jsonResponse` with the complete file.
   Preserve any unrelated functions in the project.
2. Save the project. Update its **existing web app deployment** through
   **Deploy → Manage deployments → Edit → Version: New version → Deploy**.
   Reusing the existing deployment keeps the URL configured in `index.html`.
3. Check that the `User` and `Seller` tabs have column M available. The script
   labels it `Claim ID` and writes the same ID on every item row. If column M
   already contains unrelated data, the script returns an error instead of
   overwriting it.
4. After the deployment is confirmed, merge the frontend PR. The read-only
   confirmation returns JavaScript containing this result:

   ```json
   {"confirmed":true,"claimId":"NCIG-1234567890123-ABC123","savedTo":"User","totalItems":1}
   ```

The script checks for an existing Claim ID under a script lock. Retrying the
same unchanged form in the same browser session confirms the original rows
without creating new ones.
