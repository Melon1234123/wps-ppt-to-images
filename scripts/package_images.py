from pathlib import Path
import zipfile,re,posixpath,io,hashlib,json
from lxml import etree as E
from PIL import Image
import argparse,sys
ap=argparse.ArgumentParser(description='Package WPS-rendered pages as a PPTX while preserving video slides verbatim.')
ap.add_argument('source',type=Path); ap.add_argument('images',type=Path); ap.add_argument('output',type=Path)
ap.add_argument('--max-mb',type=float,default=100); ap.add_argument('--quality',type=int,default=92)
args=ap.parse_args(); src=args.source.resolve(); out=args.output.resolve()
if src==out or out.exists(): raise SystemExit('Output must be a new file, distinct from source.')
if not 1<=args.quality<=95 or args.max_mb<=0: raise SystemExit('Invalid quality or size limit.')
z=zipfile.ZipFile(src); data={n:z.read(n) for n in z.namelist()}
P='http://schemas.openxmlformats.org/presentationml/2006/main'; A='http://schemas.openxmlformats.org/drawingml/2006/main'; R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'; PK='http://schemas.openxmlformats.org/package/2006/relationships'
ns={'p':P,'a':A,'r':R}
xml=lambda x:E.tostring(x,xml_declaration=True,encoding='UTF-8',standalone=True)
pres=E.fromstring(data['ppt/presentation.xml']); size=pres.find('p:sldSz',ns); cx,cy=size.get('cx'),size.get('cy')
for el in pres.findall('p:embeddedFontLst',ns): pres.remove(el)
data['ppt/presentation.xml']=xml(pres)
rr=E.fromstring(data['ppt/_rels/presentation.xml.rels'])
for el in list(rr):
 if el.get('Type','').endswith('/font'): rr.remove(el)
data['ppt/_rels/presentation.xml.rels']=xml(rr)
slide_paths=[n for n in data if re.fullmatch(r'ppt/slides/slide\d+.xml',n)]
video_slides=set()
for path in slide_paths:
 rp=posixpath.join(posixpath.dirname(path),'_rels',posixpath.basename(path)+'.rels')
 if rp in data:
  rel=E.fromstring(data[rp])
  if any(e.get('Type','').endswith(('/video','/media','/audio')) for e in rel):video_slides.add(path)
# Verify the full export before creating any output.
image_files={}
for path in slide_paths:
 if path in video_slides:continue
 i=int(re.search(r'slide(\d+)',path).group(1))
 matches=[p for p in args.images.iterdir() if p.suffix.lower() in ('.png','.jpg','.jpeg') and re.search(r'(\d+)$',p.stem) and int(re.search(r'(\d+)$',p.stem).group(1))==i]
 if len(matches)!=1:raise SystemExit(f'Expected exactly one WPS image for slide {i}, found {len(matches)}')
 image_files[i]=matches[0]
for path in slide_paths:
 if path in video_slides:continue
 i=int(re.search(r'slide(\d+)',path).group(1))
 path=f'ppt/slides/slide{i}.xml'; s=E.fromstring(data[path]); tree=s.find('p:cSld/p:spTree',ns)
 for el in list(tree)[2:]: tree.remove(el)
 for tag in ['timing','transition','clrMapOvr']:
  for el in s.findall('p:'+tag,ns):
   if tag!='clrMapOvr': s.remove(el)
 pic=E.fromstring(f'''<p:pic xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:nvPicPr><p:cNvPr id="2" name="WPS渲染页面{i}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdFlattened"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>''');tree.append(pic)
 data[path]=xml(s)
 rp=f'ppt/slides/_rels/slide{i}.xml.rels'; rel=E.fromstring(data[rp])
 for el in list(rel):
  if not el.get('Type','').endswith(('/slideLayout','/notesSlide')):rel.remove(el)
 E.SubElement(rel,'{'+PK+'}Relationship',Id='rIdFlattened',Type=R+'/image',Target=f'../media/flat{i}.jpg');data[rp]=xml(rel)
 im=Image.open(image_files[i]).convert('RGB');b=io.BytesIO();im.save(b,format='JPEG',quality=args.quality,optimize=True,subsampling=0);data[f'ppt/media/flat{i}.jpg']=b.getvalue()
# Retain only package parts reachable through relationships.
seen=set();todo=[''];keep={'[Content_Types].xml'}
while todo:
 part=todo.pop()
 if part in seen:continue
 seen.add(part)
 if part:keep.add(part)
 rp=posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels') if part else '_rels/.rels'
 if rp not in data:continue
 keep.add(rp)
 for rel in E.fromstring(data[rp]):
  if rel.get('TargetMode')=='External':continue
  target=posixpath.normpath(posixpath.join(posixpath.dirname(part),rel.get('Target',''))).lstrip('/')
  if target in data:todo.append(target)
ct=E.fromstring(data['[Content_Types].xml'])
for el in list(ct):
 if el.tag.endswith('Override') and el.get('PartName','').lstrip('/') not in keep:ct.remove(el)
if not any(el.get('Extension')=='jpg' for el in ct):E.SubElement(ct,'{http://schemas.openxmlformats.org/package/2006/content-types}Default',Extension='jpg',ContentType='image/jpeg')
data['[Content_Types].xml']=xml(ct)
def write_package(destination):
 with zipfile.ZipFile(destination,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as oz:
  for n in sorted(keep):oz.writestr(n,data[n])
temp=out.with_suffix('.building.pptx')
if temp.exists():raise SystemExit(f'Temporary output exists: {temp}')
try:
 write_package(temp)
 if temp.stat().st_size>=args.max_mb*1_000_000:
  for quality,width in [(85,2560),(80,1920),(72,1600),(60,1280)]:
   for n in keep:
    if re.fullmatch(r'ppt/media/flat\d+.jpg',n):
     i=int(re.search(r'flat(\d+)',n).group(1));im=Image.open(image_files[i]).convert('RGB')
     if im.width>width:im.thumbnail((width,round(width*im.height/im.width)))
     b=io.BytesIO();im.save(b,format='JPEG',quality=min(quality,args.quality),optimize=True,subsampling=0);data[n]=b.getvalue()
   write_package(temp)
   if temp.stat().st_size<args.max_mb*1_000_000:break
 if temp.stat().st_size>=args.max_mb*1_000_000:raise RuntimeError('Cannot meet size limit with unchanged media; do not recompress video without permission.')
 with zipfile.ZipFile(temp) as oz:
  assert oz.testzip() is None
  assert len([n for n in oz.namelist() if re.fullmatch(r'ppt/slides/slide\d+.xml',n)])==len(slide_paths)
  for path in video_slides:
   assert oz.read(path)==z.read(path)
   rp=posixpath.join(posixpath.dirname(path),'_rels',posixpath.basename(path)+'.rels')
   assert oz.read(rp)==z.read(rp)
  media=[]
  for n in keep:
   if n.lower().endswith(('.mp4','.avi','.wmv','.mov','.m4v','.mp3','.wav','.m4a')):
    assert oz.read(n)==z.read(n);media.append({'part':n,'sha256':hashlib.sha256(oz.read(n)).hexdigest()})
 temp.rename(out)
 print(json.dumps({'output':str(out),'bytes':out.stat().st_size,'MB':round(out.stat().st_size/1e6,2),'slides':len(slide_paths),'preserved_media_slides':sorted(video_slides),'media':media},ensure_ascii=False))
finally:
 if temp.exists():temp.unlink()
