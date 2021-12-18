# GenshinPullTracker
Google Appscript to import Genshin Impact Wish data into a Google Spreadsheet

Requires `ImportJSON.gs` from https://github.com/bradjasper/ImportJSON

To Use:
1. Create a new Spreadsheet in Google Docs.
2. Go to `Extensions` > `Apps Script`.
3. Click the `+` sign next to `Files` and choose `Script`.
4. Name the file `ImportGenshinWishes.gs`.
5. Copy the contents of `ImportGenshinWishes.gs` from this GitHub repository and paste it into the file you created above.
6. Save the file.
7. Click the `+` sign next to `Files` and choose `Script`.
8. Name the file `ImportJSON.gs`.
9. Copy the contents of `ImportJSON.gs` from https://github.com/bradjasper/ImportJSON and paste it into the file you created above.
10. Save the file.
11. Return to the spreadsheet and reload the page. A menu called `Genshin Wishes` should appear next to the Help menu.
12. Go to `Genshin Wishes` > `Fetch Latest`. Agree to all the scary messages.
13. Go to `Genshin Wishes` > `Fetch` Latest again.

Obtaining the Genshin Impact URL
1. Run Genshin Impact and open the Wish History page.
2. In your file explorer go to the `%appdata%\..\LocalLow\miHoYo\Genshin Impact` folder.
3. Open the `output_log.txt` file in a text editor.
4. Search for `OnGetWebViewPageFinish` and copy the URL (https://webstatic-sea.mihoyo.com/hk4e/event.../log).

