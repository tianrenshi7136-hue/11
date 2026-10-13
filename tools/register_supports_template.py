"""Register the engineer-selected Blinding source after template review."""
import argparse,json
from pathlib import Path
from openpyxl import load_workbook
from els.templates import sha256
from els.supports import ROOT,BLINDING_FILE

def main():
 parser=argparse.ArgumentParser();parser.add_argument('--source',type=Path,default=BLINDING_FILE);args=parser.parse_args()
 source=args.source.resolve();book=load_workbook(source,read_only=True,data_only=False,keep_links=False)
 try:
  sheet=book['Blinding'];formulas={a:sheet[a].value for a in ('D13','D18','C22','C26','G15','G16')}
  if any(not isinstance(f,str) or not f.startswith('=') for f in formulas.values()):raise ValueError('Missing Blinding formulas')
 finally:book.close()
 config={'schema_version':1,'blinding_file':str(source),'blinding_sha256':sha256(source),'expected_formulas':formulas,'approved_source':'Engineer-selected Blinding worksheet; controlled input and comparison repairs'}
 destination=ROOT/'templates/supports-structure.local.json';destination.write_text(json.dumps(config,ensure_ascii=False,indent=2),encoding='utf-8');print(destination)
if __name__=='__main__':main()
