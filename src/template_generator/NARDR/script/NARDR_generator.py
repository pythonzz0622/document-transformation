import os
import numpy as np
import pandas as pd
import re


class Convert:
    '''
    Description :
        실무자 block , 요청 block , 붙임 block 크게 3가지 block을 구축할 수 있게 구현

    '''

    def __init__(self, data_path):
        df = pd.read_excel(data_path, header=None)
        df.columns = ['key', 'value']
        self.all_rowCnt = 1

        congress = df[df['key'] == '국회의원']['value'].values[0]
        self.xml_frame = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<hs:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
        xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">
    <hp:p id="2757524817" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
        <hp:run charPrIDRef="7">
            <hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000"
                      tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0"
                      masterPageCnt="0">
                <hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>
                <hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>
                <hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL"
                               fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>
                <hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>
                <hp:pagePr landscape="WIDELY" width="59528" height="84186" gutterType="LEFT_ONLY">
                    <hp:margin header="4252" footer="4252" gutter="0" left="8504" right="8504" top="5668"
                               bottom="4252"/>
                </hp:pagePr>
                <hp:footNotePr>
                    <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>
                    <hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>
                    <hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>
                    <hp:numbering type="CONTINUOUS" newNum="1"/>
                    <hp:placement place="EACH_COLUMN" beneathText="0"/>
                </hp:footNotePr>
                <hp:endNotePr>
                    <hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>
                    <hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/>
                    <hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>
                    <hp:numbering type="CONTINUOUS" newNum="1"/>
                    <hp:placement place="END_OF_DOCUMENT" beneathText="0"/>
                </hp:endNotePr>
                <hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0"
                                   fillArea="PAPER">
                    <hp:offset left="1417" right="1417" top="1417" bottom="1417"/>
                </hp:pageBorderFill>
                <hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0"
                                   fillArea="PAPER">
                    <hp:offset left="1417" right="1417" top="1417" bottom="1417"/>
                </hp:pageBorderFill>
                <hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0"
                                   fillArea="PAPER">
                    <hp:offset left="1417" right="1417" top="1417" bottom="1417"/>
                </hp:pageBorderFill>
            </hp:secPr>
            <hp:ctrl>
                <hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/>
            </hp:ctrl>
        </hp:run>
        <hp:run charPrIDRef="7">
            <hp:t>{congress} 의원님께</hp:t>
        </hp:run>
        <hp:linesegarray>
            <hp:lineseg textpos="0" vertpos="0" vertsize="1800" textheight="1800" baseline="1530" spacing="1080"
                        horzpos="0" horzsize="42520" flags="393216"/>
        </hp:linesegarray>
    </hp:p>
        '''
        self.doc_close = '''            </hp:tbl>
            <hp:t/>
        </hp:run>
        <hp:linesegarray>
            <hp:lineseg textpos="0" vertpos="2880" vertsize="38944" textheight="38944" baseline="33102" spacing="1080"
                        horzpos="0" horzsize="42520" flags="393216"/>
        </hp:linesegarray>
    </hp:p>
    <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
        <hp:run charPrIDRef="7"/>
        <hp:linesegarray>
            <hp:lineseg textpos="0" vertpos="42904" vertsize="1800" textheight="1800" baseline="1530" spacing="1080"
                        horzpos="0" horzsize="42520" flags="393216"/>
        </hp:linesegarray>
    </hp:p>
</hs:sec>'''

    def get_tbl_frame(self):
        xml_code = f'''    <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                <hp:run charPrIDRef="7">
                    <hp:tbl id="2063134606" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES"
                            lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="{self.all_rowCnt}" colCnt="5" cellSpacing="0"
                            borderFillIDRef="3" noAdjust="0">
                        <hp:sz width="40820" widthRelTo="ABSOLUTE" height="38378" heightRelTo="ABSOLUTE" protect="0"/>
                        <hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0"
                                vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0"
                                horzOffset="0"/>
                        <hp:outMargin left="283" right="283" top="283" bottom="283"/>
                        <hp:inMargin left="510" right="510" top="141" bottom="141"/>
                        <hp:cellzoneList>
                        </hp:cellzoneList>
        '''
        return xml_code

    # 실무자 blcok
    def get_person_xml(self, df_person):

        # 실무자 수
        person_id = df_person['key'].apply(lambda x: re.sub(r'[^0-9]', '', x))
        person_id = sorted(list(set(person_id)))
        rowspan = len(person_id) + 1

        # 실무자 block xml code 뼈대
        xml_code = f'''                <hp:tr>
                    <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="3">
                        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                    hasNumRef="0">
                            <hp:p id="0" paraPrIDRef="19" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                <hp:run charPrIDRef="6">
                                    <hp:t>실무자</hp:t>
                                </hp:run>
                                <hp:linesegarray>
                                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                spacing="600" horzpos="0" horzsize="4940" flags="393216"/>
                                </hp:linesegarray>
                            </hp:p>
                        </hp:subList>
                        <hp:cellAddr colAddr="0" rowAddr="0"/>
                        <hp:cellSpan colSpan="1" rowSpan="{rowspan}"/>
                        <hp:cellSz width="5960" height="{1565 * rowspan}"/>
                        <hp:cellMargin left="510" right="510" top="141" bottom="141"/>
                    </hp:tc>
                    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">
                        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                    hasNumRef="0">
                            <hp:p id="0" paraPrIDRef="19" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                <hp:run charPrIDRef="6">
                                    <hp:t>소속</hp:t>
                                </hp:run>
                                <hp:linesegarray>
                                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                spacing="600" horzpos="0" horzsize="18240" flags="393216"/>
                                </hp:linesegarray>
                            </hp:p>
                        </hp:subList>
                        <hp:cellAddr colAddr="1" rowAddr="0"/>
                        <hp:cellSpan colSpan="2" rowSpan="1"/>
                        <hp:cellSz width="19261" height="1565"/>
                        <hp:cellMargin left="510" right="510" top="510" bottom="510"/>
                    </hp:tc>
                    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">
                        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                    hasNumRef="0">
                            <hp:p id="0" paraPrIDRef="19" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                <hp:run charPrIDRef="6">
                                    <hp:t>이름</hp:t>
                                </hp:run>
                                <hp:linesegarray>
                                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                spacing="600" horzpos="0" horzsize="6636" flags="393216"/>
                                </hp:linesegarray>
                            </hp:p>
                        </hp:subList>
                        <hp:cellAddr colAddr="3" rowAddr="0"/>
                        <hp:cellSpan colSpan="1" rowSpan="1"/>
                        <hp:cellSz width="7658" height="1565"/>
                        <hp:cellMargin left="510" right="510" top="510" bottom="510"/>
                    </hp:tc>
                    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">
                        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                    hasNumRef="0">
                            <hp:p id="0" paraPrIDRef="19" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                <hp:run charPrIDRef="6">
                                    <hp:t>전화번호</hp:t>
                                </hp:run>
                                <hp:linesegarray>
                                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                spacing="600" horzpos="0" horzsize="6920" flags="393216"/>
                                </hp:linesegarray>
                            </hp:p>
                        </hp:subList>
                        <hp:cellAddr colAddr="4" rowAddr="0"/>
                        <hp:cellSpan colSpan="1" rowSpan="1"/>
                        <hp:cellSz width="7941" height="1565"/>
                        <hp:cellMargin left="510" right="510" top="510" bottom="510"/>
                    </hp:tc>
                </hp:tr>'''

        # 실무자 number 별로 iter
        for _id in person_id:
            df = df_person[df_person['key'].str.contains(_id)]

            depart = df[df['key'].str.contains('소속')]['value'].values[0]
            name = df[df['key'].str.contains('이름')]['value'].values[0]
            number = df[df['key'].str.contains('전화번호')]['value'].values[0]
            # make 실무자 xml code
            xml_code += self.get_person_info(depart, name, number)

        return xml_code

    # 요청 block
    def get_request_xml(self, df_request):

        xml_code = ''
        # request 숫자값 추출
        df_num = df_request[df_request['key'].str.contains('요청')]['key'].str.extract(r'(\d)')

        # 중복 제거
        number_list = list(set(df_num.to_dict()[0].values()))

        # 요청 별로 요청 detail xml code 구성
        for number in number_list:
            df = df_request[df_request['key'].str.contains(f'요청{number}')]
            request_number = df['key'].values[0]
            request_sentence = df['value'].values[0]

            # add 요청 xml code block
            xml_code += self.get_requs_info(request_number, request_sentence)

            # get detail code block
            df_detail = df_request[df_request['key'].str.contains(f'{number}-')]

            # 요청 세부사항이 있을 경우
            if df_detail.empty == False:

                # add 요청 detail xml code block
                xml_code += self.get_req_detail_info(df_detail, number)

            # 요청에 바로 답변이 달릴 경우
            else:
                pass

        return xml_code


    # 붙임 block
    def get_attached_xml(self, df_attached):
        attach = df_attached['value'].values[0]
        xml_code = self.get_attached_info(attach)
        return xml_code


    # 요청 xml code return
    def get_requs_info(self, request_number, request_sentence):
        xml_code = f'''
                        <hp:tr dtype="request">
                            <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3" dtype="request_id">
                                <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                            linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                            hasNumRef="0">
                                    <hp:p id="0" paraPrIDRef="19" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                        <hp:run charPrIDRef="6">
                                            <hp:t>{request_number}</hp:t>
                                        </hp:run>
                                        <hp:linesegarray>
                                            <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                        spacing="600" horzpos="0" horzsize="4940" flags="393216"/>
                                        </hp:linesegarray>
                                    </hp:p>
                                </hp:subList>
                                <hp:cellAddr colAddr="0" rowAddr="{self.all_rowCnt}"/>
                                <hp:cellSpan colSpan="1" rowSpan="1"/>
                                <hp:cellSz width="5960" height="2416"/>
                                <hp:cellMargin left="510" right="510" top="425" bottom="425"/>
                            </hp:tc>
                            <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3" dtype="request_answer">
                                <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                            linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                            hasNumRef="0">
                                    <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                        <hp:run charPrIDRef="6">
                                            <hp:t>{request_sentence}</hp:t>
                                        </hp:run>
                                        <hp:linesegarray>
                                            <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                        spacing="600" horzpos="0" horzsize="33840" flags="393216"/>
                                        </hp:linesegarray>
                                    </hp:p>
                                </hp:subList>
                                <hp:cellAddr colAddr="1" rowAddr="{self.all_rowCnt}"/>
                                <hp:cellSpan colSpan="4" rowSpan="1"/>
                                <hp:cellSz width="34859" height="2416"/>
                                <hp:cellMargin left="510" right="510" top="425" bottom="425"/>
                            </hp:tc>
                        </hp:tr>'''
        self.all_rowCnt += 1
        return xml_code

    # 요청 세부사항 xml code return
    def get_req_detail_info(self, df_detail, number):

        df_detail = df_detail.reset_index()
        df_detail.drop('index', inplace=True, axis=1)
        rowspan = len(df_detail)

        # 뼈대 잡기
        xml_code = f'''
                <hp:tr dtype="detail_request">
                    <hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="3">
                        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                    hasNumRef="0">
                            <hp:p id="0" paraPrIDRef="19" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                <hp:run charPrIDRef="6">
                                    <hp:t>요청 {number}</hp:t>
                                </hp:run>
                                <hp:linesegarray>
                                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                spacing="600" horzpos="0" horzsize="4940" flags="393216"/>
                                </hp:linesegarray>
                            </hp:p>
                            <hp:p id="0" paraPrIDRef="19" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                <hp:run charPrIDRef="6">
                                    <hp:t>세부 사항</hp:t>
                                </hp:run>
                                <hp:linesegarray>
                                    <hp:lineseg textpos="0" vertpos="1600" vertsize="1000" textheight="1000"
                                                baseline="850" spacing="600" horzpos="0" horzsize="4940"
                                                flags="393216"/>
                                </hp:linesegarray>
                            </hp:p>
                        </hp:subList>
                        <hp:cellAddr colAddr="0" rowAddr="{self.all_rowCnt}"/>
                        <hp:cellSpan colSpan="1" rowSpan="{rowspan}"/>
                        <hp:cellSz width="5960" height="18407"/>
                        <hp:cellMargin left="510" right="510" top="141" bottom="141"/>
                    </hp:tc>
                    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">
                        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                    hasNumRef="0">
                            <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                <hp:run charPrIDRef="6">
                                    <hp:t>{df_detail.iloc[0].key}</hp:t>
                                </hp:run>
                                <hp:linesegarray>
                                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                spacing="600" horzpos="0" horzsize="4372" flags="393216"/>
                                </hp:linesegarray>
                            </hp:p>
                        </hp:subList>
                        <hp:cellAddr colAddr="1" rowAddr="{self.all_rowCnt}"/>
                        <hp:cellSpan colSpan="1" rowSpan="1"/>
                        <hp:cellSz width="5394" height="5614"/>
                        <hp:cellMargin left="510" right="510" top="1417" bottom="1417"/>
                    </hp:tc>
                    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">
                        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                    hasNumRef="0">
                            <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                <hp:run charPrIDRef="6">
                                    <hp:t>{df_detail.iloc[0].value}
                                    </hp:t>
                                </hp:run>
                                <hp:linesegarray>
                                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                spacing="600" horzpos="0" horzsize="26632" flags="393216"/>
                                    <hp:lineseg textpos="31" vertpos="1600" vertsize="1000" textheight="1000"
                                                baseline="850" spacing="600" horzpos="0" horzsize="26632"
                                                flags="393216"/>
                                </hp:linesegarray>
                            </hp:p>
                        </hp:subList>
                        <hp:cellAddr colAddr="2" rowAddr="{self.all_rowCnt}"/>
                        <hp:cellSpan colSpan="3" rowSpan="1"/>
                        <hp:cellSz width="29465" height="5614"/>
                        <hp:cellMargin left="1417" right="1417" top="1417" bottom="1417"/>
                    </hp:tc>
                </hp:tr>'''
        self.all_rowCnt += 1

        # 세부 사항 별 iter
        for df_i in df_detail[1:].iterrows():
            xml_code += f'''
                            <hp:tr dtype="detail_request">
                    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">
                        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                    hasNumRef="0">
                            <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                <hp:run charPrIDRef="6">
                                    <hp:t>{df_i[1].key}</hp:t>
                                </hp:run>
                                <hp:linesegarray>
                                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                spacing="600" horzpos="0" horzsize="4372" flags="393216"/>
                                </hp:linesegarray>
                            </hp:p>
                        </hp:subList>
                        <hp:cellAddr colAddr="1" rowAddr="{self.all_rowCnt}"/>
                        <hp:cellSpan colSpan="1" rowSpan="1"/>
                        <hp:cellSz width="5394" height="5614"/>
                        <hp:cellMargin left="510" right="510" top="1417" bottom="1417"/>
                    </hp:tc>
                    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">
                        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                    hasNumRef="0">
                            <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                <hp:run charPrIDRef="6">
                                    <hp:t>{df_i[1].value}
                                    </hp:t>
                                </hp:run>
                                <hp:linesegarray>
                                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                spacing="600" horzpos="0" horzsize="26632" flags="393216"/>
                                    <hp:lineseg textpos="32" vertpos="1600" vertsize="1000" textheight="1000"
                                                baseline="850" spacing="600" horzpos="0" horzsize="26632"
                                                flags="393216"/>
                                    <hp:lineseg textpos="62" vertpos="3200" vertsize="1000" textheight="1000"
                                                baseline="850" spacing="600" horzpos="0" horzsize="26632"
                                                flags="393216"/>
                                </hp:linesegarray>
                            </hp:p>
                        </hp:subList>
                        <hp:cellAddr colAddr="2" rowAddr="{self.all_rowCnt}"/>
                        <hp:cellSpan colSpan="3" rowSpan="1"/>
                        <hp:cellSz width="29465" height="5614"/>
                        <hp:cellMargin left="1417" right="1417" top="1417" bottom="1417"/>
                    </hp:tc>
                </hp:tr>'''
            self.all_rowCnt += 1
        return xml_code

    # 실무자 xml code return
    def get_person_info(self, depart, name, number):
        xml_code = f'''<hp:tr dtype="manager">
    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3" dtype="depart">
        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                    hasNumRef="0">
            <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                <hp:run charPrIDRef="6">
                    <hp:t>{depart}</hp:t>
                </hp:run>
                <hp:linesegarray>
                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                spacing="600" horzpos="0" horzsize="18240" flags="393216"/>
                </hp:linesegarray>
            </hp:p>
        </hp:subList>
        <hp:cellAddr colAddr="1" rowAddr="{self.all_rowCnt}"/>
        <hp:cellSpan colSpan="2" rowSpan="1"/>
        <hp:cellSz width="19261" height="1565"/>
        <hp:cellMargin left="510" right="510" top="510" bottom="510"/>
    </hp:tc>
    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3" dtype="name">
        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                    hasNumRef="0">
            <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                <hp:run charPrIDRef="6">
                    <hp:t>{name}</hp:t>
                </hp:run>
                <hp:linesegarray>
                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                spacing="600" horzpos="0" horzsize="6636" flags="393216"/>
                </hp:linesegarray>
            </hp:p>
        </hp:subList>
        <hp:cellAddr colAddr="3" rowAddr="{self.all_rowCnt}"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:cellSz width="7658" height="1565"/>
        <hp:cellMargin left="510" right="510" top="510" bottom="510"/>
    </hp:tc>
    <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3" dtype="number">
        <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                    linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                    hasNumRef="0">
            <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                <hp:run charPrIDRef="6">
                    <hp:t>{number}</hp:t>
                </hp:run>
                <hp:linesegarray>
                    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                spacing="600" horzpos="0" horzsize="6920" flags="393216"/>
                </hp:linesegarray>
            </hp:p>
        </hp:subList>
        <hp:cellAddr colAddr="4" rowAddr="{self.all_rowCnt}"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:cellSz width="7941" height="1565"/>
        <hp:cellMargin left="510" right="510" top="510" bottom="510"/>
    </hp:tc>
    </hp:tr>'''

        self.all_rowCnt += 1
        return xml_code

    # 붙임 xml code return
    def get_attached_info(self, attach):
        xml_code = f'''
                        <hp:tr>
                        <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3">
                            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                        linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                        hasNumRef="0">
                                <hp:p id="0" paraPrIDRef="19" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                    <hp:run charPrIDRef="6">
                                        <hp:t>붙임</hp:t>
                                    </hp:run>
                                    <hp:linesegarray>
                                        <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                    spacing="600" horzpos="0" horzsize="4940" flags="393216"/>
                                    </hp:linesegarray>
                                </hp:p>
                            </hp:subList>
                            <hp:cellAddr colAddr="0" rowAddr="{self.all_rowCnt}"/>
                            <hp:cellSpan colSpan="1" rowSpan="1"/>
                            <hp:cellSz width="5960" height="2416"/>
                            <hp:cellMargin left="510" right="510" top="425" bottom="425"/>
                        </hp:tc>
                        <hp:tc name="" header="0" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="3" dtype="attach">
                            <hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"
                                        linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0"
                                        hasNumRef="0">
                                <hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">
                                    <hp:run charPrIDRef="6">
                                        <hp:t>{attach}</hp:t>
                                    </hp:run>
                                    <hp:linesegarray>
                                        <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850"
                                                    spacing="600" horzpos="0" horzsize="33840" flags="393216"/>
                                    </hp:linesegarray>
                                </hp:p>
                            </hp:subList>
                            <hp:cellAddr colAddr="1" rowAddr="{self.all_rowCnt}"/>
                            <hp:cellSpan colSpan="4" rowSpan="1"/>
                            <hp:cellSz width="34859" height="2416"/>
                            <hp:cellMargin left="510" right="510" top="425" bottom="425"/>
                        </hp:tc>
                    </hp:tr>'''
        self.all_rowCnt += 1
        return xml_code


if __name__ == "__main__":

    # data 불러오기 및 class 불러오기
    convert = Convert('../data/sample_label_3.xlsx')
    df = pd.read_excel('../data/sample_label_3.xlsx', header=None)

    # xml code block 을 만드는데 필요한 dataframe 나누기
    df.columns = ['key', 'value']
    df_answer = df.loc[df['key'] == '전체답변']
    df_person = df[df['key'].str.contains('실무자')]
    df_attached = df.loc[df['key'] == '붙임']
    df_request = df[df['key'].str.contains('요청') | df['key'].str.contains('답변')]

    ################## xml code block 쌓기 ##################################
    full_xml_code_list = [convert.xml_frame]

    # 실무자 block
    person_xml = convert.get_person_xml(df_person)
    full_xml_code_list.append(person_xml)

    # 요청 block
    request_xml = convert.get_request_xml(df_request)
    full_xml_code_list.append(request_xml)

    # 붙임 block
    attach_xml = convert.get_attached_xml(df_attached)
    full_xml_code_list.append(attach_xml)

    # xml 끝 부분 block
    full_xml_code_list.append(convert.doc_close)

    # xml tbl 앞부분 (block cell 개수를 계산해야 되기 때문에 맨 마지막에 불러오기)
    full_xml_code_list.insert(1, convert.get_tbl_frame())

    # xml code 붙이기
    xml_code = ''
    for code in full_xml_code_list:
        xml_code += code
    full_xml_code = xml_code.replace('\n', '')

    # xml code 저장
    with open(f'../output/Contents/section0.xml', 'w') as f:
        f.write(full_xml_code)
    print('done')

    # hwpx파일 생성
    os.chdir('../output')
    os.system('zip -0 -r ../sample.hwpx $(ls)')
