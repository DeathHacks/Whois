

import requests
import pandas as pd
import time
import sys
from tldextract import extract
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed
from ratelimit import limits, sleep_and_retry


@sleep_and_retry
# Limits 50 request per 1 second 
@limits(calls=50, period=1) 
#Function to send request. 
def send_request(domain):
    attempts = 0
    max_retries = 50
    decode_attempt = 0
    apikey = "KEY"
    #API request information 
    url = "https://www.whoisxmlapi.com/whoisserver/WhoisService?apiKey=%s&outputFormat=JSON&domainName=%s" %apikey, domain
    while attempts < max_retries:
            try:
                #Send request 
                response = requests.request("GET", url)
                if str(response.status_code).startswith('5'):
                    response.raise_for_status()
                #return response to main function without formating, handled in main function 
                return response
            except Exception or requests.exceptions.HTTPError as e:
                #On error, add to attempt variable and retry after a second
                print("error, retrying", response.json(),e)
                attempts +=1
                print("Take a break")
                time.sleep(5) 
    #After 10 attempts, quit program
    if attempts or decode_attempt == max_retries: 
        print("API has dropped 10 requests consecutively."'/n'" Make sure API isn't down."'/n',e )
        export_results()


def cleandomain (domain):
    #Extract domain, using extract means no need for regex
    domain_extract = extract(domain)
    #Split subdomain value  on . 
    domain_split = domain_extract.subdomain.split('.')
    #Check amount of dots in subdomain if 1 or less return value. 
    if len(domain_split) <= 1:
        return domain
    else:
        #If more then 1 , return the last value always forth level domain, six.fifth.SPLIT.third.second.first Example this.is.a.domain.com.au
        domain = domain_split[-1] + "."+domain_extract.domain+"."+domain_extract.suffix
        return domain

def process_domain(domain):
    #Assign original domain for future use in keypair list
    original_domain = domain["Domain"]
    #Submit domain to clean domain function 
    clean_domain = cleandomain(domain["Domain"])
    #Send returned clean_domain to send request function
    response = send_request(clean_domain)
    #Get json data from response and assign result
    result = response.json()
    try:
        Owner = result["WhoisRecord"]["registryData"]["registrant"]["organization"]
    except: 
        Owner = "N/A"
    if Owner == "N/A":
        return {"Domain": original_domain, "Owner":"N/A", "Status": "Inactive"}
    elif Owner is None:
        return {"Domain": original_domain, "Owner": "N/A", "Status": "Active"}
    else:
        return {"Domain": original_domain, "Owner": Owner, "Status": "Active"}

def export_results():
    #Close statusbar entries 
    pbar.close()
    print("WHOIS extraction complete, writing to file")
    df_output = pd.DataFrame(data)
    # Save the dataframe to an Excel file with headers "Domain" and "Owner"
    df_output.to_excel("H:/DomainOwners.xlsx", index=False)
    print("Write to file complete, check C:\temp")
    sys.exit()


if __name__ == "__main__":
    # Load the spreadsheet into a Pandas dataframe
    df = pd.read_excel("H:/Domainlist.xlsx")
    print(df.columns)
    data = []
    i = 0
    workers = 50
    print("Checking WHOIS data for ",len(df)," domains.")
    #Progress bar define 
    pbar = tqdm(total=len(df), desc="Processing WHOIS Data")
    # Define ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=workers) as executor:
        #Submits domain field to process_domain for each domain in df
        futures = [executor.submit(process_domain, domain) for _, domain in df.iterrows()]
        #Ensure no duplicate futures 
        for future in as_completed(futures):
            result = future.result()
            #append to data array 
            data.append(result)
            #Update status bar 
            pbar.update(1)
            i+=1
            if i == len(df):
                export_results()