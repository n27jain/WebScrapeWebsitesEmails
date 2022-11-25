import re
import requests
from urllib.parse import urlsplit
from collections import deque
from bs4 import BeautifulSoup
import pandas as pd
from google.colab import files
 
# read url from input
original_url = input("cleanearthinc.com") 
 
# to save urls to be scraped
unscraped = deque([original_url])
 
# to save scraped urls
scraped = set()
 
# to save fetched emails
emails = set()  
 
while len(unscraped):
    print("go")
    url = unscraped.popleft()  
    scraped.add(url)
 
    parts = urlsplit(url)
        
    base_url = "{0.scheme}://{0.netloc}".format(parts)
    if '/' in parts.path:
      path = url[:url.rfind('/')+1]
    else:
      path = url
 
    print("Crawling URL %s" % url)
    try:
        response = requests.get(url)
    except (requests.exceptions.MissingSchema, requests.exceptions.ConnectionError):
        continue
 
    new_emails = set(re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.com", response.text, re.I))
    emails.update(new_emails) 
 
    soup = BeautifulSoup(response.text, 'lxml')
 
    for anchor in soup.find_all("a"):
      if "href" in anchor.attrs:
        link = anchor.attrs["href"]
      else:
        link = ''
 
        if link.startswith('/'):
            link = base_url + link
        
        elif not link.startswith('http'):
            link = path + link

print(emails)