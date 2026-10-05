"""Dependency-free, website-only Pages artifact builder. Never reads release payloads."""
from pathlib import Path
import argparse, json, os, re, shutil, xml.etree.ElementTree as ET
from urllib.parse import urlsplit
from html import escape

site = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser()
parser.add_argument('--site-url', default=os.environ.get('SITE_URL', ''))
args = parser.parse_args()
config = (site/'site-config.js').read_text(encoding='utf-8')
def field(key):
 match = re.search(r'\b'+key+r':\s*"([^"]*)"',config)
 return match[1] if match else ''
version = ET.parse(site.parent/'Directory.Build.props').findtext('.//ProductVersion')
assert re.fullmatch(r'\d+\.\d+\.\d+(?:[-+][\w.-]+)?',version or ''), 'Invalid product version'
config = re.sub(r'(productVersion:\s*)"[^"]*"',lambda m:m[1]+json.dumps(version),config)
domain = field('customDomain')
if domain and not re.fullmatch(r'[a-zA-Z0-9](?:[a-zA-Z0-9.-]*[a-zA-Z0-9])?',domain):
 raise SystemExit('customDomain must contain only a hostname')
url = field('siteUrl') or ('https://'+domain+'/' if domain else args.site_url)
if url:
 parts = urlsplit(url)
 if parts.scheme != 'https' or not parts.hostname or parts.query or parts.fragment or parts.username or parts.password:
  raise SystemExit('siteUrl must be an absolute HTTPS root with optional project path')
 url = url.rstrip('/')+'/'
 config = re.sub(r'(siteUrl:\s*)"[^"]*"',lambda m:m[1]+json.dumps(url),config)
output = site/'_build'
assert output.resolve().parent == site.resolve() and output.name == '_build'
if output.exists(): shutil.rmtree(output) # Only this fixed, generated website artifact.
output.mkdir()
# Deliberate publish allowlist excludes scripts, documents, tests, manifests and binaries.
for folder in ('assets','css','js'):
 shutil.copytree(site/folder,output/folder)
(output/'site-config.js').write_text(config,encoding='utf-8')
pages = []
for source in sorted(site.rglob('index.html')):
 if '_build' in source.parts: continue
 relative = source.relative_to(site)
 text = source.read_text(encoding='utf-8')
 text = re.sub(r'(<span data-version>)[^<]+',lambda m:m[1]+version,text)
 if url:
  canonical = url + (relative.parent.as_posix()+'/' if relative.parent != Path('.') else '')
  description = re.search(r'<meta name="description" content="([^"]*)"',text)[1]
  title = re.search(r'<title>(.*?)</title>',text)[1]
  image = url+'assets/brand/social.png'
  metadata = f'<link rel="canonical" href="{escape(canonical,quote=True)}"><meta property="og:url" content="{escape(canonical,quote=True)}"><meta property="og:image" content="{escape(image,quote=True)}"><meta name="twitter:image" content="{escape(image,quote=True)}">'
  data = {'@context':'https://schema.org','@type':'WebPage','name':title,'url':canonical,'description':description}
  metadata += '<script type="application/ld+json">'+json.dumps(data,ensure_ascii=False).replace('<','\\u003c')+'</script>'
  text = text.replace('</head>',metadata+'</head>')
  pages.append(canonical)
 target = output/relative; target.parent.mkdir(parents=True,exist_ok=True);target.write_text(text,encoding='utf-8')
robots = 'User-agent: *\nAllow: /\n'
if url:
 robots += 'Sitemap: '+url+'sitemap.xml\n'
 sitemap = ET.Element('urlset',xmlns='http://www.sitemaps.org/schemas/sitemap/0.9')
 for page in pages: ET.SubElement(ET.SubElement(sitemap,'url'),'loc').text = page
 ET.ElementTree(sitemap).write(output/'sitemap.xml',encoding='utf-8',xml_declaration=True)
(output/'robots.txt').write_text(robots,encoding='utf-8')
(output/'.nojekyll').touch()
if domain:(output/'CNAME').write_text(domain+'\n',encoding='utf-8')
print(f'Built {len(list(output.rglob("index.html")))} pages; version {version}; canonical root: {url or "not configured"}')
