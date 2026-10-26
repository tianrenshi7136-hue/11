"""Optional PDF regression check (Poppler, pdfplumber and Pillow)."""
import argparse,json,subprocess,tempfile
from pathlib import Path
import pdfplumber
import numpy as np
from PIL import Image,ImageChops,ImageStat,ImageDraw


def main():
 p=argparse.ArgumentParser();p.add_argument('result');p.add_argument('--baseline',default='outputs/waling-weld-correction/20261023-172651-eb1729/result.json');p.add_argument('--allow-report-date',action='store_true');args=p.parse_args()
 before=json.loads(Path(args.baseline).read_text(encoding='utf8'));after=json.loads(Path(args.result).read_text(encoding='utf8'))
 folder=Path(after['pdf']).parent;checks=[]
 with tempfile.TemporaryDirectory(prefix='pdf-compare-',dir=folder) as temp, pdfplumber.open(before['pdf']) as a,pdfplumber.open(after['pdf']) as b:
  for label,path in [('before',before['pdf']),('after',after['pdf'])]:
   subprocess.run(['pdftoppm','-r','108','-png',path,str(Path(temp)/label)],check=True,capture_output=True,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
  assert len(b.pages)==len(after['output_sheets']),'Page count changed'
  indices=[before['output_sheets'].index(name) for name in after['output_sheets']]
  for number,(index,y) in enumerate(zip(indices,b.pages),1):
   x=a.pages[index]
   assert (x.width,x.height)==(y.width,y.height),(number,'Page size changed')
   before_text,after_text=x.extract_text(),y.extract_text();masked=[]
   if args.allow_report_date:
    import re
    pattern=r'\b\d\s?\d-[A-Z][a-z]{2}-\d{2}\b'
    old=re.findall(pattern,before_text);new=re.findall(pattern,after_text)
    assert len(old)==len(new)==1,(number,'Report date is ambiguous')
    before_text=before_text.replace(old[0],'REPORT DATE');after_text=after_text.replace(new[0],'REPORT DATE')
    masked=[(w['x0'],w['top'],w['x1'],w['bottom']) for page in (x,y) for w in page.search(pattern)]
    assert len(masked)==2,(number,'Report date bounds missing')
   assert before_text==after_text,(number,'Printed text changed')
   def geometry(page):
    return sorted((kind,*(round(item[k],3) for k in ('x0','top','x1','bottom','linewidth'))) for kind,items in [('line',page.lines),('rect',page.rects)] for item in items)
   same_geometry=geometry(x)==geometry(y)
   def extent(page):
    items=page.lines+page.rects
    return tuple(round(fn(item[k] for item in items),3) for fn,k in ((min,'x0'),(min,'top'),(max,'x1'),(max,'bottom')))
   assert extent(x)==extent(y),(number,'Frame position changed')
   first=Image.open(Path(temp)/f'before-{index+1}.png').convert('RGB');second=Image.open(Path(temp)/f'after-{number}.png').convert('RGB')
   if number in (1,5):second.save(folder/f'page-{number}.png')
   for rect in masked:
    bounds=(int(rect[0]*1.5)-1,int(rect[1]*1.5)-1,int(rect[2]*1.5)+2,int(rect[3]*1.5)+2)
    for im in (first,second):ImageDraw.Draw(im).rectangle(bounds,fill='white')
   diff=ImageChops.difference(first,second);bounds=diff.getbbox();mean=ImageStat.Stat(diff).mean
   changed=float(np.any(np.asarray(diff)!=0,axis=2).mean())
   checks.append(dict(page=number,identical_vector_segments=same_geometry,identical_pixels=bounds is None,changed_pixel_fraction=changed,mean_channel_difference=mean,difference_bounds=bounds))
   assert changed<.002 and max(mean)<.02,(number,'PDF appearance changed',bounds,mean)
 report=dict(status='PASS',pages=len(checks),report_date_excluded=args.allow_report_date,checks=checks,verified='Same text, frame extent and page size; vector segmentation recorded; raster variation below 0.2% pixels and 0.02/255 mean intensity at 108 dpi')
 (folder/'pdf-comparison.json').write_text(json.dumps(report,indent=2),encoding='utf8')
 print(json.dumps(report),flush=True)


if __name__=='__main__':main()
