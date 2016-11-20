
# coding: utf-8

# In[ ]:

import time
import requests
import webbrowser
import os
from lxml import html
import tempfile
import pandas as pd
from datetime import date


# In[ ]:

def browse(page):
    """
        :param page:
            Pass in the the requests element
            page = requests.get(url)
        :return: current HTML.
            browse(page)
    """
    tmpdir = tempfile.mkdtemp()
    filename = 'browse.html'
    path = os.path.join(tmpdir, filename)
    file_path = 'file://' + path
    print path
    # writing file
    with open(path, "w") as f:
        try:
            f.write(page.encode('ascii','ignore'))
        except:
            f.write(page.text.encode('ascii','ignore'))
        f.close()
    return webbrowser.open(file_path)


# In[ ]:

def getProductInfo(page_html, prod_id):
    today = date.today().isoformat()
    product = {
            'brand': ''
            ,'short_description': ''
            ,'price': ''
            ,'color': ''
            ,'picture_url': ''
            ,'product_id': prod_id
            ,'store': 'Saks_Fifth_Avenue'
            ,'date': today
        } 

    try:
        brand = page_html.xpath('//*[@id="'+prod_id+'"]/div[3]/a/p[1]/span')[0].text
        try:
            product['brand'] = brand.strip()
        except AttributeError:
            pass
    except IndexError:
        pass
    
    try:
        short_description = page_html.xpath('//*[@id="'+prod_id+'"]/div[3]/a/p[2]')[0].text
        try:
            product['short_description'] = short_description.strip()
        except AttributeError:
            pass
    except IndexError:
        pass
    
    try:
        price = page_html.xpath('//*[@id="'+prod_id+'"]/div[3]/a/span')[0].text
        try:
            product['price'] = price.strip()
        except AttributeError:
            pass
    except IndexError:
        pass

    try:
        product['color'] = page_html.xpath('//*[@id="'+prod_id+'"]/div[1]/div')[0].get('data-label')
    except IndexError:
        pass
        
    try:
        product['picture_url'] = page_html.xpath('//*[@id="'+prod_id+'"]/div[1]/div/img')[0].get('data-src')
    except IndexError:
        pass
    
    return product


# In[ ]:

headers = {
    'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
    ,'Accept-Encoding':'gzip, deflate, sdch'
    ,'Accept-Language':'en-US,en;q=0.8'
    ,'Cache-Control':'max-age=0'
    ,'Connection':'keep-alive'
    ,'Host':'www.saksfifthavenue.com'
    ,'Pragma':'no-cache'
    ,'Upgrade-Insecure-Requests':'1'
    ,'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.86 Safari/537.36'
}
session = requests.session()
page = session.get('http://www.saksfifthavenue.com/Entry.jsp', headers=headers)
page_html = html.fromstring(page.text)
url = page_html.xpath('//*[@id="WomensApparelNavMenu"]/ul/li[1]/ul/li[1]/a')[0].attrib['href']
headers.update({
    'Referer':'http://www.saksfifthavenue.com/Entry.jsp'
})

page = session.get(url, headers=headers)
page_html = html.fromstring(page.text)

df = pd.DataFrame()
page_nums = page_html.xpath('//*[@id="pc-top"]/div[3]/span[2]')[0].text.strip()
for i in range(1, int(page_nums)):
    print i
    for product in page_html.xpath('//*[@id="product-container"]')[0]:
        if product.get('id') != None:
            prod_id = product.get('id')
            df = df.append(pd.DataFrame(getProductInfo(page_html, prod_id),index=[prod_id]))
    
    headers.update({ 'Referer': page.url })
    url = 'http://www.saksfifthavenue.com%s' % page_html.xpath('//*[@id="pc-top"]/ol/li[8]/a')[0].get('href')        
    page = session.get(url, headers=headers)
    page_html = html.fromstring(page.text)

    


# In[ ]:

df.to_excel('saks_pricing_data_20160502.xlsx')


# In[ ]:

browse(page)


# In[ ]:



