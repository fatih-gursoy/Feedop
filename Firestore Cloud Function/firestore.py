import firebase_admin

from firebase_admin import credentials
from firebase_admin import firestore
import requests
from bs4 import BeautifulSoup as bs
from datetime import datetime
from time import strptime
import lxml

Haber_Listesi = dict()
Tarih, Haber, Link  = [],[],[]

def RSSHaber_cek(rss_url):
   url = rss_url
   r = requests.get(url)
   soup = bs(r.content, features ='xml')
   items = soup.find_all("item")

   for item in items:
      uzun_tarih = str(item.pubDate.text).split(" ")
      uzun_tarih[2] = strptime(uzun_tarih[2],"%b").tm_mon
      trh = ("{}.{}.{}".format(uzun_tarih[1], uzun_tarih[2], uzun_tarih[3]))
      
      Tarih.append(trh)
      Haber.append(item.title.text)
      Link.append(item.link.text)

      Haber_Listesi["Haber"] = Haber
      Haber_Listesi["Link"] = Link
      Haber_Listesi["Tarih"] = Tarih

   return Haber_Listesi

url_list = ["https://www.marketingturkiye.com.tr/feed/", 
"https://www.campaigntr.com/category/haberler/feed/", 
"https://webrazzi.com/feed/"]

for url in url_list:
   Haber_Listesi = RSSHaber_cek(url)

cred = credentials.Certificate("./serviceAccountKey.json")
default_app = firebase_admin.initialize_app(cred)
db = firestore.client()

kontrol = {"Haber":[],"Link":[], "Tarih":[]}
docid, docs =[], []

for hbr in Haber_Listesi["Haber"]:
    doc = db.collection('Haberler_Listesi').where('Haber', 'array_contains', hbr).get()
    
    for item in doc:
      if not item.id in docid:
          docs.append(item)
          docid.append(item.id)
      
for doc in docs:
    for k, v in doc.to_dict().items():
        kontrol[k] += v

liste1 = list(zip(Haber_Listesi["Haber"], Haber_Listesi["Link"], Haber_Listesi["Tarih"]))
liste2 = list(zip(kontrol["Haber"], kontrol["Link"], kontrol["Tarih"]))

now = datetime.now().strftime("%d%m%Y_%H%M%S")

if len(liste2) == 0:
    db.collection("Haberler_Listesi").document(now).set(Haber_Listesi)
else:
    yeni_haberler =[]
    [yeni_haberler.append(i) for i in liste1 if not i in liste2]
    
    if len(yeni_haberler) != 0:
        filter_haber = {"Haber":[],"Link":[], "Tarih":[]}

        for a,b,c in yeni_haberler:
            filter_haber["Haber"].append(a)
            filter_haber["Link"].append(b)
            filter_haber["Tarih"].append(c)
        
        db.collection("Haberler_Listesi").document(now).set(filter_haber)

    