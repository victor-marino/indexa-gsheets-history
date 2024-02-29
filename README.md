# Description
You can use this Google Apps Script to automate the process of fetching your daily account positions from Indexa Capital and writing them to a Google Sheet spreadsheet in your own Google account.

Your account data is accessed securely by using Indexa's [official API](https://support.indexacapital.com/es/esp/introduccion-api)

This can be useful for:
* Keeping an offline history of the value of your portfolio in your own spreadsheets, without having to access your Indexa Capital account or pull out any exports.
* Build the price history of instruments that do not have a publicly available NAV, such as Indexa's own pension plans for self-employed people.

# Usage

## Setting up the script
1. Go to script.google.com and start a new project.
3. In the code editor, delete the existing code and paste everything from the [code.gs](https://github.com/victor-marino/indexa-gsheets-history/blob/master/code.gs) file found in this repository.
4. Replace `YOUR_API_TOKEN` with your Indexa Capital API token. Instructions on how to obtain it [here](https://support.indexacapital.com/es/esp/introduccion-api).
5. Replace `YOUR_ACCOUNT_NUMBER` with the number of the account you want to monitor. It is an 8-character alphanumeric code that you can see in the web/app when browsing your account.
6. Choose a file name for the spreadsheet where the data will be stored. If doesn't exist, it will be created. If it exists, sheets will be created inside it to store the data.
7. If you want the resulting tables to be sorted from newest to oldest, set `newestOnTop` to `true`. This will ensure the latest value is always at the top.
8. Give a name to your project and save it. Can be anything you want.

Time for a test run!

## Testing the script
Above the code editor, make sure the `run` function is selected:

![google_app_script_run](https://github.com/victor-marino/indexa-gsheets-history/assets/1933443/86dce857-5998-492b-80a9-057e3d2f0379)

Now hit the `Run` button:

![Captura de pantalla 2024-02-29 a las 19 17 59](https://github.com/victor-marino/indexa-gsheets-history/assets/1933443/8c10257e-ffe9-4eba-abf3-7054e1d2b1db)

Since this is the first run, the script will prompt you for all the required permissions to work with your Google Drive files. Once you grant them, the script should go ahead and create a new spreadsheet containing one sheet per instrument in your account (e.g.: one sheet per index fund). If the result looks like this, all went well:

![Captura de pantalla 2024-02-29 a las 19 34 33](https://github.com/victor-marino/indexa-gsheets-history/assets/1933443/50ae6fac-312d-4116-a65d-e6d93b4df70a)

If something went wrong, you can take a look at the error messages in the log to figure out what happened.

## Scheduling runs

Now all that's left is to program the script so that it runs everyday on its own, without any human intervention.

To do this, just head over to `Triggers` in the left side menu, then click the `+ Add Trigger` button:

![Captura de pantalla 2024-02-29 a las 19 39 36](https://github.com/victor-marino/indexa-gsheets-history/assets/1933443/3b4113a3-962f-43d3-aaf7-28d66fda04ab)

* Function to run: `run`
* What to run: `Main`
* Event source: `Time-driven`
* Type of time based trigger: `Day timer`
* Time of day: `12:00 to 13:00` (or whatever you prefer)

![Captura de pantalla 2024-02-29 a las 19 45 35](https://github.com/victor-marino/indexa-gsheets-history/assets/1933443/255a014f-f785-4fe3-9578-50227f8fb9d0)

And that's it! Press `Save` and you're all set.

The script should run daily from now on, updating all your positions in the spreadsheet if there is new data.

Remember mutual funds and pension plans do not update their prices in bank holidays and other special dates, so some days the script will have nothing to add to the table.
