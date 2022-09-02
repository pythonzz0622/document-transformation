# code structure

```
src
 |-parser 
   |- NDD : html templates 에서 csv로 [필요 정보]()를 추출합니다.
   |- NANDR : xml(hwpx) templates 에서 csv로 [필요 정보]()를 추출합니다.

 |-templates_generator 
   |- NDD : csv 을 html template 로 변환합니다.
   |- NANDR : csv 을 hwpx template 로 변환합니다.
    
``` 

## NDD generator
```
# start mogodb
mongod --dbpath=/Users/mac/data/db

# flask run
python -m flask run
```

## NARDR generator

```
python NARDR_generator.py 
```

NARDR 문제점 
1. 페이지가 넘어갈경우
2. input 문장의 길이와 lineseg의 tag가 일치하지 않는 경우 