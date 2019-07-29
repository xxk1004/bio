import time
import xlwt
import requests
from bs4 import BeautifulSoup

keyword = "MDM2 breast cancer"
url = "https://www.ncbi.nlm.nih.gov/pubmed/?term="
url = url + keyword
total_pages = 31


def fetch_item_page(url):
    content = requests.get(url).content
    soup = BeautifulSoup(content)
    rprt_div_tag = soup.find(name='div', attrs={"class": "rprt_all"})
    # print(rprt_div_tag)
    title = rprt_div_tag.find(name='h1').get_text()
    print(title)
    try:
        abstract = rprt_div_tag.find(name='div', attrs={"class": "abstr"}).find(name='p').get_text()
    except:
        abstract = ''
    # print(abstract)
    try:
        doi = rprt_div_tag.find(name='dl', attrs={"class": "rprtid"}).find_all(name='dd')[1].get_text()
    except:
        doi = ''
    # print(doi)
    time.sleep(1)
    return title, abstract, doi


def crawl_page(session, url, data):
    headers = {
        "origin": "https://www.ncbi.nlm.nih.gov",
        "referer": "https://www.ncbi.nlm.nih.gov/pubmed",
        # "cookie": "_ga=GA1.2.672535571.1564282185; _gid=GA1.2.36961716.1564282185; ncbi_sid=277188115AA16696_62C4SID; QSI_HistorySession=https%3A%2F%2Fwww.ncbi.nlm.nih.gov%2F~1564282186226; WebEnv=1JESjMkRW-HaC0g90zj-JmdYfX6l8PTAe7xavSnwXzgkxo56Ese39wZkb0e4qL4oHPJo48yZLrOhID1KXQ0VSpeMmPcBSDMh5KnQZ%40277188115AA16696_62C4SID; _gat_ncbiSg=1; ncbi_pinger=N4IgDgTgpgbg+mAFgSwCYgFwgMIFEAc2AbPgEy4AiAzBQCy7akCMADG++wOykCCLVAMQoA6ALZwmtEABoQAYwA2yOQGsAdlAAeAF0ygWmcAEMA5sjVHtyAPZrpV7Qqhwja1AlNQI0ube1Q1bR9rBWkAZ1gAuF8FGRAmQ1Ejc19AgKDcQOgALwFrCFE4qgBOQ1pOFhIqBNlaAyxyySZ8GpBaKSxFZXUtXVqAVkMDWqJDADMjBQi48sN+lnwZxax+zlLa0qwjOSsYKGkwTzglNRVpDR04+cNEbW0wMIwAeieAd3fhNTkAI2RPhVEn2QiGEJmsMCeYAArt9RFBUABiK4JLCZbQ5ADKAE8wv5RKRhAAFNE5IkwgCy8LJsPhcAASlAwlCFNowoTXFAFNS4e52SYvETPFdSIYAHK9AAEAD4QABfWRQtQKaxGVAXXQYUBUKiGdFQqBFDogPUG2RUZYgIqbNoVEizWr1G1NYqcGZGxXK1XqmaDLCtfqjBqcf2urDDEBEFHxFicQby+TWUSiWzezUgEVYdI5IqGaE09C1QyoaxyJmFWS+y2yQMgYulqHlkChkC3USxWQW0jh62oKATZl9eJRkWyJg6rCR9bxI2HfnRRCuflxJg1ybt+LNiZTU3xC1b6ayLuGPCEEjkah0BjMDg37h8QQicSSOLMQxZqDZDB5nkYOtljCigA8qKuAvhmIDvK8nw/H8SqAmowKguCL7WkwxRHmajqsPg4bVIYaEYSAVDgawqwzOOIAsMIpD9MI/oUZwRA6hWRr4NcsqykAA",
        "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36"
    }
    content = session.post(url=url, data=data, headers=headers).content
    # print(content)
    soup = BeautifulSoup(content)
    div_list = soup.find_all(name='div', attrs={"class": "rprt"})
    return_list = []
    for div_tag in div_list:
        p_tag = div_tag.find(name='p')
        a_tag = p_tag.find(name='a')
        item_url = "https://www.ncbi.nlm.nih.gov" + a_tag['href']
        title, abstract, doi = fetch_item_page(item_url)
        row = {
            "title": title,
            "abstract": abstract,
            "doi": doi
        }
        return_list.append(row)
        # print(title)
        # print(abstract)
        # print(doi)
        # break
        # print(item_url)
    return return_list


wbk = xlwt.Workbook()
sheet = wbk.add_sheet('Sheet1', cell_overwrite_ok=True)
row_num = 0

session = requests.session()
content = session.get(url).content
# print(content)
soup = BeautifulSoup(content)
div_list = soup.find_all(name='div', attrs={"class": "rprt"})
for div_tag in div_list:
    p_tag = div_tag.find(name='p')
    a_tag = p_tag.find(name='a')
    item_url = "https://www.ncbi.nlm.nih.gov" + a_tag['href']
    title, abstract, doi = fetch_item_page(item_url)
    sheet.write(row_num, 0, title)
    sheet.write(row_num, 1, abstract)
    sheet.write(row_num, 2, doi)
    row_num = row_num + 1
    # print(title)
    # break
    # print(item_url)
for page_num in range(2, total_pages + 1):
    print("Current page: " + str(page_num))
    post_dict = {
        "term": keyword,
        "EntrezSystem2.PEntrez.PubMed.Pubmed_PageController.PreviousPageName": "results",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_PageController.SpecialPageName": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_Facets.FacetsUrlFrag": "filters=",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_Facets.FacetSubmitted": "false",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_Facets.BMFacets": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.sPresentation": "docsum",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.sSort": "none",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.sPageSize": "20",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FFormat": "csv",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FSort": "",
        "email_format": "docsum",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.email_sort": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.email_count": "20",
        "email_start": "1",
        "email_address": "",
        "email_subj": "",
        "email_add_text": "",
        "EmailCheck1": "",
        "EmailCheck2": "",
        "BibliographyUser": "",
        "BibliographyUserName": "my",
        "citman_count": "20",
        "citman_start": "1",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FileFormat": "csv",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastPresentation": "docsum",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.Presentation": "docsum",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PageSize": "20",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastPageSize": "20",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.Sort": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastSort": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FileSort": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.Format": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastFormat": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PrevPageSize": "20",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PrevPresentation": "docsum",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PrevSort": "",
        "CollectionStartIndex": "1",
        "CitationManagerStartIndex": "1",
        "CitationManagerCustomRange": "false",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_ResultsController.ResultCount": "619",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_ResultsController.RunLastQuery": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.cPage": "1",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.CurrPage": str(page_num),
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.cPage": "1",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailReport": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailFormat": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailCount": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailStart": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailSort": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Email": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailSubject": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailText": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailQueryKey": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailHID": "1H7u2UoO0BM3Ueweldy-YSBnBWKsKAMlFMDPUPFT3UylmBWsMGT80BJ4CSVpYIN9VF1HQK0i_ZuIoiXByr_RIMQHP3By1zpc4D",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.QueryDescription": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Key": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Answer": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Holding": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.HoldingFft": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.HoldingNdiSet": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.OToolValue": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.SubjectList": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.TimelineAdPlaceHolder.CurrTimelineYear": "",
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.TimelineAdPlaceHolder.BlobID": "NCID_1_266684448_130.14.22.33_9001_1564283695_1121658376_0MetA0_S_MegaStore_F_1",
        "EntrezSystem2.PEntrez.DbConnector.Db": "pubmed",
        "EntrezSystem2.PEntrez.DbConnector.LastDb": "pubmed",
        "EntrezSystem2.PEntrez.DbConnector.Term": keyword,
        "EntrezSystem2.PEntrez.DbConnector.LastTabCmd": "",
        "EntrezSystem2.PEntrez.DbConnector.LastQueryKey": "1",
        "EntrezSystem2.PEntrez.DbConnector.IdsFromResult": "",
        "EntrezSystem2.PEntrez.DbConnector.LastIdsFromResult": "",
        "EntrezSystem2.PEntrez.DbConnector.LinkName": "",
        "EntrezSystem2.PEntrez.DbConnector.LinkReadableName": "",
        "EntrezSystem2.PEntrez.DbConnector.LinkSrcDb": "",
        "EntrezSystem2.PEntrez.DbConnector.Cmd": "PageChanged",
        "EntrezSystem2.PEntrez.DbConnector.TabCmd": "",
        "EntrezSystem2.PEntrez.DbConnector.QueryKey": "",
        "p$a": "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.Page",
        "p$l": "EntrezSystem2",
        "p$st": "pubmed"
    }
    # post_param = "term=MDM2+breast+cancer+&EntrezSystem2.PEntrez.PubMed.Pubmed_PageController.PreviousPageName=results&EntrezSystem2.PEntrez.PubMed.Pubmed_PageController.SpecialPageName=&EntrezSystem2.PEntrez.PubMed.Pubmed_Facets.FacetsUrlFrag=filters%3D&EntrezSystem2.PEntrez.PubMed.Pubmed_Facets.FacetSubmitted=false&EntrezSystem2.PEntrez.PubMed.Pubmed_Facets.BMFacets=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.sPresentation=docsum&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.sSort=none&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.sPageSize=20&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FFormat=csv&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FSort=&email_format=docsum&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.email_sort=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.email_count=20&email_start=1&email_address=&email_subj=MDM2+breast+cancer+-+PubMed&email_add_text=&EmailCheck1=&EmailCheck2=&BibliographyUser=&BibliographyUserName=my&citman_count=20&citman_start=1&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FileFormat=csv&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastPresentation=docsum&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.Presentation=docsum&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PageSize=20&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastPageSize=20&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.Sort=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastSort=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FileSort=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.Format=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastFormat=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PrevPageSize=20&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PrevPresentation=docsum&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PrevSort=&CollectionStartIndex=1&CitationManagerStartIndex=1&CitationManagerCustomRange=false&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_ResultsController.ResultCount=619&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_ResultsController.RunLastQuery=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.cPage=1&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.CurrPage=2&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.cPage=1&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailReport=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailFormat=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailCount=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailStart=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailSort=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Email=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailSubject=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailText=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailQueryKey=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailHID=1H7u2UoO0BM3Ueweldy-YSBnBWKsKAMlFMDPUPFT3UylmBWsMGT80BJ4CSVpYIN9VF1HQK0i_ZuIoiXByr_RIMQHP3By1zpc4D&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.QueryDescription=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Key=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Answer=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Holding=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.HoldingFft=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.HoldingNdiSet=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.OToolValue=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.SubjectList=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.TimelineAdPlaceHolder.CurrTimelineYear=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.TimelineAdPlaceHolder.BlobID=NCID_1_266684448_130.14.22.33_9001_1564283695_1121658376_0MetA0_S_MegaStore_F_1&EntrezSystem2.PEntrez.DbConnector.Db=pubmed&EntrezSystem2.PEntrez.DbConnector.LastDb=pubmed&EntrezSystem2.PEntrez.DbConnector.Term=MDM2+breast+cancer&EntrezSystem2.PEntrez.DbConnector.LastTabCmd=&EntrezSystem2.PEntrez.DbConnector.LastQueryKey=1&EntrezSystem2.PEntrez.DbConnector.IdsFromResult=&EntrezSystem2.PEntrez.DbConnector.LastIdsFromResult=&EntrezSystem2.PEntrez.DbConnector.LinkName=&EntrezSystem2.PEntrez.DbConnector.LinkReadableName=&EntrezSystem2.PEntrez.DbConnector.LinkSrcDb=&EntrezSystem2.PEntrez.DbConnector.Cmd=PageChanged&EntrezSystem2.PEntrez.DbConnector.TabCmd=&EntrezSystem2.PEntrez.DbConnector.QueryKey=&p%24a=EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.Page&p%24l=EntrezSystem2&p%24st=pubmed"
    result_list = crawl_page(session=session, url="https://www.ncbi.nlm.nih.gov/pubmed", data=post_dict)
    for row in result_list:
        sheet.write(row_num, 0, row['title'])
        sheet.write(row_num, 1, row['abstract'])
        sheet.write(row_num, 2, row['doi'])
        row_num = row_num + 1
    time.sleep(2)

wbk.save('MDM2_breast_cancer.xls')
