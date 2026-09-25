# Warranty form Apps Script update

The browser form sends `claimId` with each claim. The current deployed Apps Script
does not save or return this ID, so the stricter browser confirmation must not
be published until this script is deployed.

1. In the existing warranty Apps Script project, replace only `doPost` and
   `jsonResponse` with the functions in [Code.gs](Code.gs). Preserve any other
   functions in the project.
2. Save the project. Update its **existing web app deployment** through
   **Deploy → Manage deployments → Edit → Version: New version → Deploy**.
   Reusing the existing deployment keeps the URL configured in `index.html`.
3. Check that the `User` and `Seller` tabs have column M available. The script
   labels it `Claim ID` and writes the same ID on every item row. If column M
   already contains unrelated data, the script returns an error instead of
   overwriting it.
4. After the deployment is confirmed, merge the frontend PR. A successful
   response has this shape:

   ```json
   {"success":true,"claimId":"NCIG-1234567890123-ABC123","savedTo":"User","totalItems":1}
   ```

The script checks for an existing Claim ID under a script lock. Retrying the
same unchanged form in the same browser session confirms the original rows
without creating new ones.
