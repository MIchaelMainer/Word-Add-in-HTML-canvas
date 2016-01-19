MUST have the reg key noted here : https://github.com/OfficeDev/Office-Add-in-Commands-Samples

TODO: note that add-in commands won't work on the Mac or iPad.
TODO: Create my own map.


tsd install
npm install
gulp copy:libs
gulp


For Mac (note that the add-in commands won't work)
1.	Create a folder called “wef” in Users/<username>/Library/Containers/com.microsoft.word/Data/Documents/
2.	Put the developer manifest in the wef folder (Users/<username>/Library/Containers/com.microsoft.word/Data/Documents/wef)
3.	Open word application on Mac and click on inset->”my add-ins” drop down.
