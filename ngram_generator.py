try:
	import re
	import os
	import sys
	import numpy as np
	import pandas as pd
	from sklearn.feature_extraction.text import CountVectorizer
except:
	print('do pip install -r requirement.txt')
	exit()
	

parenth=re.compile('\(\w*\)')
othr=re.compile('((jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|sun|mon|tue|wed|thr|fri|sat|to:|from:|re:)( |,))')
email=re.compile('\S*@\S*\s?')
date=re.compile('(?:(?:[0-9]{1,2}[:\/,]){1,2}[0-9]{2,4}|\s+am|\s+pm)')
newline=re.compile('\n')
punc=re.compile('[^a-z\s]')
multispace=re.compile('[ ]+')


def transform(x):
  x=x.lower()
  x=parenth.sub('',x)
  x=email.sub('',x)
  x=date.sub('',x)
  x=newline.sub(' ',x)
  x=x.replace('''\xa0''',' ')
  x=othr.sub('',x)
  x=punc.sub('',x)
  x=multispace.sub(' ',x.strip())
  return x



def main():
	print('Will crawl in current directory and its subdir for xlsx and xls file to generate ngram. \n')
	col=input('Enter Column name only one column \n')
	a=int(input('Enter ngram start range 2 for bigram etc \n'))
	b=int(input('Enter ngram end range \n'))
	for root, subdirs, files in os.walk(os.path.abspath('.')):
		for fname in files:
			if fname.endswith('.xlsx') or fname.endswith('.xls'):
				print(f'Working with {fname}')
				#starts here

				opname='ngram_'+fname
				with pd.ExcelWriter(root+'/'+opname) as writer:  
					for n in range(a,b+1):
						print(f'Generating {str(n)}gram ')
						ip=pd.read_excel(root+'/'+fname)
						df = pd.DataFrame(ip[col].apply(transform))
						word_vectorizer = CountVectorizer(ngram_range=(n,n), analyzer='word')
						sparse_matrix = word_vectorizer.fit_transform(df[col])
						frequencies = sum(sparse_matrix).toarray()[0]
						(pd.DataFrame(frequencies, index=word_vectorizer.get_feature_names(), columns=['frequency'])).sort_values(['frequency'],ascending=False).to_excel(
							writer,
							sheet_name=str(n)+'gram'
						)

				#ends here

if __name__ == '__main__':
		main()