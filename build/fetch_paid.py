import json, os, subprocess, urllib.request
from concurrent.futures import ThreadPoolExecutor
ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
jobs=json.load(open(os.path.join(ROOT,'build','paid_photo_jobs.json')))
UA='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36'
def one(j):
    url,out,w=j; out=os.path.join(ROOT,out)
    if os.path.exists(out): return 'skip'
    try:
        u=url.split('=')[0]+'=w800-h600-k-no' if 'googleusercontent' in url else url
        data=urllib.request.urlopen(urllib.request.Request(u,headers={'User-Agent':UA}),timeout=30).read()
        tmp=out+'.tmp'; open(tmp,'wb').write(data)
        subprocess.run(['sips','-s','format','jpeg','-s','formatOptions','58','--resampleWidth',str(w),tmp,'--out',out],check=True,capture_output=True)
        os.remove(tmp); return 'ok'
    except Exception as ex: return 'ERR '+str(ex)[:60]
with ThreadPoolExecutor(8) as ex: res=list(ex.map(one,jobs))
print('paid photos ok',sum(r=='ok' for r in res),'err',sum(r.startswith('ERR') for r in res))
