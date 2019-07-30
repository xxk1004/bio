import time
import xlwt
import requests
from bs4 import BeautifulSoup
import parameters

keyword = "MDM2 breast cancer"
url = "https://www.ncbi.nlm.nih.gov/pubmed/?term="
url = url + keyword
total_pages = 31


def fetch_item_page(url):
    headers = {
        "origin": "https://www.ncbi.nlm.nih.gov",
        "referer": "https://www.ncbi.nlm.nih.gov/pubmed",
        "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36"
    }
    content = requests.get(url=url, headers=headers).content
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
    time.sleep(2)
    return title, abstract, doi


def crawl_page(session, url, data):
    headers = {
        "origin": "https://www.ncbi.nlm.nih.gov",
        "referer": "https://www.ncbi.nlm.nih.gov/pubmed",
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
        try:
            title, abstract, doi = fetch_item_page(item_url)
        except:
            print("Retrying...")
            time.sleep(5)
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


session = requests.session()
content = session.get(url).content
# print(content)
soup = BeautifulSoup(content)
div_list = soup.find_all(name='div', attrs={"class": "rprt"})


'''wbk = xlwt.Workbook()
sheet = wbk.add_sheet('page 1', cell_overwrite_ok=True)
row_num = 0
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
wbk.save('MDM2_breast_cancer.xls')'''


for page_num in range(2, total_pages + 1):
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet('page ' + str(page_num), cell_overwrite_ok=True)
    row_num = 0
    print("Current page: " + str(page_num))
    post_dict = {
        "term": keyword,
        "EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.CurrPage": str(page_num),
        "EntrezSystem2.PEntrez.DbConnector.Term": keyword,
    }
    post_dict.update(parameters.static_post_dict)
    # post_param = "term=MDM2+breast+cancer+&EntrezSystem2.PEntrez.PubMed.Pubmed_PageController.PreviousPageName=results&EntrezSystem2.PEntrez.PubMed.Pubmed_PageController.SpecialPageName=&EntrezSystem2.PEntrez.PubMed.Pubmed_Facets.FacetsUrlFrag=filters%3D&EntrezSystem2.PEntrez.PubMed.Pubmed_Facets.FacetSubmitted=false&EntrezSystem2.PEntrez.PubMed.Pubmed_Facets.BMFacets=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.sPresentation=docsum&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.sSort=none&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.sPageSize=20&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FFormat=csv&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FSort=&email_format=docsum&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.email_sort=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.email_count=20&email_start=1&email_address=&email_subj=MDM2+breast+cancer+-+PubMed&email_add_text=&EmailCheck1=&EmailCheck2=&BibliographyUser=&BibliographyUserName=my&citman_count=20&citman_start=1&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FileFormat=csv&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastPresentation=docsum&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.Presentation=docsum&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PageSize=20&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastPageSize=20&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.Sort=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastSort=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.FileSort=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.Format=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.LastFormat=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PrevPageSize=20&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PrevPresentation=docsum&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_DisplayBar.PrevSort=&CollectionStartIndex=1&CitationManagerStartIndex=1&CitationManagerCustomRange=false&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_ResultsController.ResultCount=619&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_ResultsController.RunLastQuery=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.cPage=1&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.CurrPage=2&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.cPage=1&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailReport=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailFormat=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailCount=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailStart=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailSort=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Email=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailSubject=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailText=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailQueryKey=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.EmailHID=1H7u2UoO0BM3Ueweldy-YSBnBWKsKAMlFMDPUPFT3UylmBWsMGT80BJ4CSVpYIN9VF1HQK0i_ZuIoiXByr_RIMQHP3By1zpc4D&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.QueryDescription=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Key=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Answer=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.Holding=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.HoldingFft=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.HoldingNdiSet=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.OToolValue=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.EmailTab.SubjectList=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.TimelineAdPlaceHolder.CurrTimelineYear=&EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.TimelineAdPlaceHolder.BlobID=NCID_1_266684448_130.14.22.33_9001_1564283695_1121658376_0MetA0_S_MegaStore_F_1&EntrezSystem2.PEntrez.DbConnector.Db=pubmed&EntrezSystem2.PEntrez.DbConnector.LastDb=pubmed&EntrezSystem2.PEntrez.DbConnector.Term=MDM2+breast+cancer&EntrezSystem2.PEntrez.DbConnector.LastTabCmd=&EntrezSystem2.PEntrez.DbConnector.LastQueryKey=1&EntrezSystem2.PEntrez.DbConnector.IdsFromResult=&EntrezSystem2.PEntrez.DbConnector.LastIdsFromResult=&EntrezSystem2.PEntrez.DbConnector.LinkName=&EntrezSystem2.PEntrez.DbConnector.LinkReadableName=&EntrezSystem2.PEntrez.DbConnector.LinkSrcDb=&EntrezSystem2.PEntrez.DbConnector.Cmd=PageChanged&EntrezSystem2.PEntrez.DbConnector.TabCmd=&EntrezSystem2.PEntrez.DbConnector.QueryKey=&p%24a=EntrezSystem2.PEntrez.PubMed.Pubmed_ResultsPanel.Pubmed_Pager.Page&p%24l=EntrezSystem2&p%24st=pubmed"
    result_list = crawl_page(session=session, url="https://www.ncbi.nlm.nih.gov/pubmed", data=post_dict)
    for row in result_list:
        sheet.write(row_num, 0, row['title'])
        sheet.write(row_num, 1, row['abstract'])
        sheet.write(row_num, 2, row['doi'])
        row_num = row_num + 1
    wbk.save('MDM2_breast_cancer.xls')
    time.sleep(2)


