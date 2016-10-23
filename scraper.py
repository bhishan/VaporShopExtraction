import re 
import json 
import mechanize
import openpyxl

url = ['https://www.namastevaporizers.com/products/arizer-solo-vaporizer']

try:
    wb = openpyxl.load_workbook("outputscrape.xlsx", read_only=False)
    ws = wb.active
except:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['url', 'sku', 'category', 'type', 'stock status', 'data-bread-crumbs', 'variants'])


re_pattern = re.compile("var meta = (.*)")
re_pattern_bread_crumbs = re.compile("data-bread-crumbs=(.*)")
re_pattern_availability = re.compile('<link itemprop="availability"(.*)')

def parse_info(url, response_content):
    for matches in re.finditer(re_pattern, response_content):
        req_body = matches.groups()
    req = req_body[0] 
    req = req[0:-1]
    json_obj = json.loads(req)
    req_json_obj = json_obj['product']
    try:
        item_type = req_json_obj['type']
        category = item_type
    except:
        item_type = ''
        category = ''
    try:
        sku = ''
        for each_variants in req_json_obj['variants']:
            #variants = req_json_obj['variants'][0]
            sku += each_variants['sku'] + ','
    except:
        sku = ''

    try:
        variants = ''
        for each_vari in req_json_obj['variants']:
            variants += each_vari['public_title'] + ','
    except:
        variants = ''

    

    try:
        for matches in re.finditer(re_pattern_bread_crumbs, response_content):
            bread_body = matches.groups()
        bread = bread_body[0]
        bread = bread[0:-1]
        bread = bread.replace('"', '')
    except:
        bread = ''
    
    try:
        for matches in re.finditer(re_pattern_availability, response_content):
            availability_body = matches.groups()
        availability = availability_body[0]
        if 'InStock' in str(availability):
            in_stock = 'In stock'
        else:
            in_stock = 'Not in stock'
    except:
        in_stock = ''
    ws.append([url, sku, category, item_type, in_stock, bread, variants])
    wb.save('outputscrape.xlsx')


def main(url):
    br = mechanize.Browser()
    br.set_handle_robots(False)
    br.addheaders = [("User-agent","Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.2.13) Gecko/20101206 Ubuntu/10.10 (maverick) Firefox/3.6.13")] 
    for each_url in url:
        try:
            resp = br.open(each_url)
            response_content = resp.read()
            parse_info(each_url, response_content)
        except:
            print "failed for url", each_url

    


if __name__ == '__main__':
    main(url)
