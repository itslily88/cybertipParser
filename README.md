# cybertipParser
Python3. Intakes json file from an ICAC Cybertip, outputs an excel file parsing out the IP addresses with ISP queries and an excel file parsing out the media.

# Requirements:
- pip install openpypl, requests
- An internet connection

# Usage
Drag and drop the .json that you obtain from IDS onto the script or .exe file. It will create two excel files (_cybertip#_Parsed_IPs.xlsx and _cybertip#_Parsed_Media.xlsx) in the same directory as the .json file.

It is HIGHLY recommended that you create a free account at arin.net and then follow the steps below to create an API key, as this will drastically decrease the time the script takes to run with CTs that have a large number of IP addresses. Once you copy and paste the API key into the script, it will create a new file 'apiKey.txt' that will store this and it wont be needed anymore.

# Setting Up an ARIN API Key
1. Create an ARIN.net account
This is pretty straight forward. Go to ARIN.net and in the upper-right hand corner click “Log In”. In the next page, choose “Create a user account” towards the top. Fill in the appropriate information, confirm your account, and login.

2. Navigate to API Key page
After creating an account and logging in, click on your account name in the upper-right hand corner to activate the drop down and select “Settings”. 
On the next page, the top box will be labeled “Security Info” with a drop down in the bottom right hand corner of that box called “Actions”. Click the “Actions” drop down and select “Manage API Keys”. This will take you to a new page to set-up your actual API key

3. Create API Key
IMPORTANT – Once you create your API key, you CANNOT retrieve it again. So take a picture, copy into a text document, do something to save it.
The top box in the window is labeled “Generate an API Key”. In the bottom right hand corner of this screen is a button called “Create API Key”. Do it.
