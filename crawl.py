import requests

keyword = "MDM2+breast+cancer"
url = "https://www.ncbi.nlm.nih.gov/pubmed/?term="
url = url + keyword

content = requests.get(url).content
print(content)