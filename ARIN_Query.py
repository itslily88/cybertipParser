#!Python3
# Created by Thomas 'Lily' Lilienthal OCT.2024 - Deschutes County, Oregon - v1
'''
Expects an IP address (ipAddress = list), and optional arin.net API key (apiKey = string).
Queries arin.net to obtain ISP names.
Returns a dictionary (output = {IPAddress:ISP}).
As of OCT.2024, arin.net allows for 10 queries per minute on a non API account
or 100 per minute on a free API account.

required library:
requests
'''
from time import sleep
import requests

def queryIP(inIPs: list, apiKey: str ='') -> dict:
    """
    Queries information about the given list of IP addresses. Can handle more faster with an API key.

    Args:
        inIPs (list): List of IP addresses to query.
        apiKey (str, optional): API key to allow for more queries within 1 minute.

    Returns:
        dict: {ipAddress : ISP name}.
    """

    # Establish required headers for arin.net query
    headers = {'Accept': 'application/json', 'Content-Type': 'application/json', 'Authorization': f'APIKEY {apiKey}'}

    # Identify sleep interval to not overload arin.net based on apiKey or not.
    if apiKey != '': pause = 90
    else: pause = 9
    current = 0
    
    # Dictionary that will eventually return info and deduped inIPs list
    output = {}
    dedupedIPs = []

    # A 'just in case' dedupe to lessen load on arin.net
    for a in inIPs: 
        if a not in dedupedIPs: dedupedIPs.append(a)

    # Make status update variable for user
    left = len(dedupedIPs)

    # Explain to the user how may IP addresses it needs to query.
    print(f'''{left} unique IP address(es) within this Cybertip. Querying arin.net for ISP information.''')

    # Loop through queries and add returns to dictionary.
    for ipAddress in dedupedIPs:
        # Establish get address and get
        url = f'https://whois.arin.net/rest/ip/{ipAddress}'
        
        # Sometimes the .get returns with a crashed connection. So this will error handle is 3 times.
        attempts = 0
        while attempts < 3:
            try:
                response = requests.get(url, headers=headers)
                attempts = 3
            except:
                print(f'Got an error with IP address {ipAddress}, trying again.')
                sleep(0.5)
                attempts += 1
                response = 'Could not connect to arin.net'
        
        # If data is returned valid, parse. If not, return error.
        if type(response) == str:
            output[ipAddress] = response
        elif response.status_code == 200:
            data = response.json()
            orgRef = data['net'].get('orgRef', {})  # Use .get() to handle missing 'orgRef'
            orgName = orgRef.get('@name', 'Org Name Not Found')  # Provide a default value
            output[ipAddress] = orgName
        else:
            return(f'Error: {response.status_code}')
        
        # de-increment status update and let the user know every 25 IP addresses where its at
        left-=1
        if left % 25 == 0 and left != 0: print(f'''{left} IP addresses to go.''')
        elif left == 0: print('Completed querying IP addresses.')

        # This goes faster than connections can be established, so we need a very small sleep between gets
        sleep(0.1)

        # Iterate and send to sleep based on if there is an apiKey passed or not to not crash connection
        current+= 1
        if current == pause: 
            print('Limit reached. Sleeping for 1 minute. The script will continue after.')
            sleep(60)
    
    # Return the completed dictionary
    return(output)