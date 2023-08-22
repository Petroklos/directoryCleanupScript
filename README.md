# Directory Cleanup Script

This script is designed to clean up specific directories based on certain criteria.

## Usage

1. Modify the `directoryPaths` array to include the paths of the directories you want to clean.

2. Set the `daysToRetain` variable to specify the number of days a folder should be retained before deletion.

3. Choose Cleaning Method:
   - **Cleaning by Date**: Uncomment the line `DirectoryCleanByDate Folder.Path` within the `DirectoryDFS` function to delete folders based on their creation date.
   - **Cleaning by Name**: Uncomment the line `DirectoryCleanByName Folder.Path` within the `DirectoryDFS` function to delete folders based on a name convention.

4. Run the script using a VBScript interpreter.

## Functions

- `DirectoryCleanByDate`: Deletes folders based on creation date.
- `DirectoryCleanByName`: Deletes folders based on a name convention.
- `DirectoryDFS`: Performs depth-first search for cleaning folders.

## Logging

The script logs its actions in the `directoryDeletion.log` file, recording user, timestamp, and deleted folders.

## Disclaimer

Use this script with caution. Make sure to test thoroughly before using it in a production environment.
