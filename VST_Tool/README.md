# VST-Tool
Modify VST settings including individual parameter details (such as decimal places, colors, etc.) and strategy device settings (shadow rates, A7 custom settings, etc.). Can also be used to build a new VST file and apply all the desired custom settings.

### When making a new version
1. Ensure that the xlsm file containing the code changes (in the ```Code``` directory) has been committed, in addition to the code changes themselves
2. Bump version of ```Const CurrentVersion``` in ```aaaWebUpdate.bas``` module
3. Update the date on the "Instructions" worksheet
4. Provide some information about the changes in the "Revision History" worksheet
5. Commit updated Excel file and version code, create a tag for the release, and make a pull request
6. Once PR is merged, update the latest version at https://www.vpsedata.ford.com/version to match ```CurrentVersion``` in the tool
7. Upload the new xlsm file to the path specified in the tool versions web app
