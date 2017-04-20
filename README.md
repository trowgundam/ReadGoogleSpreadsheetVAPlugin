# Read Google Spreadsheet VoiceAttack Plugin
This VoiceAttack plugin will attempt to read a Google Spreadsheet and return the values requested. Upon start up it will prompt the user to authorize with the Google API. This typically opens a webpage that prompts the user to Login and Authorize the plugin. This authorization is saved typically, so this should only happen the first time VoiceAttack starts with this plugin.

It accepts the following contexts from VoiceAttack:
1. `ClearData` - Deletes all data from previous runs.
1. `Authorize` - Attempts to re-authorize with Google API, same thing as is done with Init.
1. Any other value is assumed to be a Google Spreadsheet Id
   1. A variable is expected named `RequestSheetRange`. This is the range to retrieve in [A1 notation](https://developers.google.com/sheets/api/guides/concepts#a1_notation).
   1. The plugin passes back values in two ways. All values will be passed back in a variable named `SheetRange`. If there is an authorization issue, an error is returned saying such, other wise it is all data in the range 1 Row per Line, with all values in the row comma delimited.
   1. Each individual value in the Range will be passed back in a variable named `SheetRangeRow{0}Col{1}` where `{0}` is the Row Number and `{1}` is the Column Number, both 0-Based.

To install extract `ReadGoogleSheetVAPlugin.7z` to the `VoiceAttack\Apps` folder and restart/start VoiceAttack. Included is a profile with a sample command on how to use the Plugin.
