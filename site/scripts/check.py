"""Check all local HTML links, anchors, image files and publish exclusions."""
from pathlib import Path
from html.parser import HTMLParser
from urllib.parse import urlsplit, unquote
import sys
root = Path(sys.argv[1] if len(sys.argv)>1 else 'site/_build').resolve()
assert (root/'index.html').is_file(), 'Artifact must have index.html directly at its root'
assert not (root/'site').exists(), 'Artifact must not contain an enclosing site/ directory'
class Page(HTMLParser):
 def __init__(self,text):
  super().__init__(convert_charrefs=True);self.links=[];self.ids=set();self.h1=0;self.feed(text)
 def handle_starttag(self,tag,attrs):
  a=dict(attrs)
  if 'id' in a: assert a['id'] not in self.ids, 'Duplicate ID';self.ids.add(a['id'])
  if tag=='h1':self.h1+=1
  if tag=='img':assert 'alt' in a,'Missing image alt'
  for key in ('href','src'):
   if key in a:self.links.append(a[key])
pages={p:Page(p.read_text(encoding='utf-8')) for p in root.rglob('*.html')}
errors=[];count=0
for source,page in pages.items():
 if page.h1 != 1:errors.append(f'{source}: expected one h1')
 for link in page.links:
  count+=1;u=urlsplit(link)
  if u.scheme or u.netloc:continue
  if u.path.startswith('/'):errors.append(f'{source}: root-relative link {link}');continue
  dest=(source.parent/unquote(u.path)).resolve() if u.path else source
  if not dest.is_relative_to(root):errors.append(f'{source}: path escapes site: {link}');continue
  if dest.is_dir():dest=dest/'index.html'
  if not dest.exists():errors.append(f'{source}: missing {link}')
  elif u.fragment and dest in pages and unquote(u.fragment) not in pages[dest].ids:errors.append(f'{source}: missing anchor {link}')
for p in root.rglob('*'):
 if p.suffix.lower() in {'.exe','.xll','.pem','.dll','.cs','.ps1'}:errors.append(f'Forbidden artifact {p}')
 if p.is_file() and p.suffix in {'.html','.js','.json','.txt'}:
  text=p.read_text(encoding='utf-8')
  if 'C:\\' in text or 'LicenseGenerator' in text:errors.append(f'Internal path/name in {p}')
if errors:raise SystemExit('\n'.join(errors))
print(f'PASS: {len(pages)} pages, {count} link/asset references, anchors, image alt attributes, publish exclusions')
