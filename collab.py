import pandas as pd
import time
from requests_html import HTMLSession

#Excel spreadsheet to read in
excel_df = pd.read_excel(r'<full path>.xlsx')
#Fill in brand name
brand = ''
#Fill in brand name email extension e.g. @google.com
brand_email = ''
sites = excel_df.loc[:,'Site']
#Remove all reference to requests if 'Requests' column is not contained within initial spreadsheet
requests = excel_df.loc[:,'Requests']
requests_column = pd.Series(requests)

results = {}

for url in sites:
    results[f'{url}'] = {}
    try:
        if 'slack' in url:
            results[f'{url}']['Platform'] = 'Slack'
        elif 'teams' in url:
            results[f'{url}']['Platform'] = 'Teams'
        elif 'trello' in url:
            results[f'{url}']['Platform'] = 'Trello'
        else:
            results[f'{url}']['Platform'] = ''
    except:
        results[f'{url}']['Platform'] = ''
    if 'slack.com' not in url:
        results[f'{url}']['Channel'] = ''
        results[f'{url}']['Admin Invite Only'] = ''
        results[f'{url}']['Accessible with Email'] = ''
        results[f'{url}'][f'Accessible with {brand} Email Address'] = ''
        continue
    session = HTMLSession()
    print(url)
    httpsurl = "https://"+url
    try:
        r = session.get(httpsurl)
    except:
        results[f'{url}']['Channel'] = ''
        results[f'{url}']['Admin Invite Only'] = ''
        results[f'{url}']['Accessible with Email'] = ''
        results[f'{url}'][f'Accessible with {brand} Email Address'] = ''
        results[f'{url}']['Platform'] = ''
        r.close()
        session.close()
        next
    try:
        r.html.render(timeout=5, sleep=5)
        print(r.status_code)
    except:
        results[f'{url}']['Channel'] = ''
        results[f'{url}']['Admin Invite Only'] = ''
        results[f'{url}']['Accessible with Email'] = ''
        results[f'{url}'][f'Accessible with {brand} Email Address'] = ''
        results[f'{url}']['Platform'] = ''
        r.close()
        session.close()
        next
    try:
        signin_span = r.html.find('#signin_header > span', first=True)
        results[f'{url}']['Channel'] = signin_span.text
    except:
        results[f'{url}']['Channel'] = ''
    try:
        invite = r.html.find('#page_contents > div > div:nth-child(3) > p > span',first=True)
        if 'Contact the workspace administrator for an invitation' in invite.text:
            results[f'{url}']['Admin Invite Only'] = 'Yes'
        else:
            results[f'{url}']['Admin Invite Only'] = 'No'
    except:
        results[f'{url}']['Admin Invite Only'] = ''
    try:
        team_email_domains = r.html.find('#page_contents > div > div:nth-child(3) > p > strong > span',first=True)
        results[f'{url}']['Accessible with Email'] = team_email_domains.text
        if brand_email in team_email_domains.text:
            results[f'{url}'][f'Accessible with {brand} Email Address'] = 'Yes'
        else:
            results[f'{url}'][f'Accessible with {brand} Email Address'] = 'No'
    except:
        results[f'{url}']['Accessible with Email'] = ''
        results[f'{url}'][f'Accessible with {brand} Email Address'] = ''
    try:
        deleted = r.html.find('#page_contents > div.card.align_center.span_4_of_6.col.float_none.margin_auto.large_bottom_margin.right_padding > p:nth-child(2)',first=True)
        if 'deleted' in deleted.text:
            results[f'{url}']['Workspace Deleted'] = 'Yes'
        else:
            results[f'{url}']['Workspace Deleted'] = 'No'
    except:
        results[f'{url}']['Workspace Deleted'] = ''
    r.close()
    session.close()
print(results)
dataFrameObj = pd.DataFrame(results)
dfObj = dataFrameObj.transpose()
dfObj.index.name = 'URL'
dfObj['Requests'] = requests_column.values
dfObj.to_excel("Collab_Results.xlsx")
