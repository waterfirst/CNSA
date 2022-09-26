#!/usr/bin/env python
# coding: utf-8

# In[38]:


import olefile
import pandas as pd


# In[39]:


import os

#현재 폴더 경로; 작업 폴더 기준
print(os.getcwd())

import os
os.chdir("D:/Private_Documents/삼성고/")
print(os.getcwd())


# In[82]:


f = olefile.OleFileIO('자소서 질문들(최종본).hwp')
#PrvText 스트림 내의 내용을 읽기


# In[41]:


encoded_text = f.openstream("PrvText").read()
 #인코딩된 텍스트를 UTF-16으로 디코딩
decoded_text = encoded_text.decode("UTF-16")
decoded_text


# In[78]:


items = decoded_text.replace("*", "")
items = decoded_text.replace("\r", "  ")
items = items.split("\n")
list_df = pd.DataFrame(items)


# In[77]:


list_df


# In[52]:


import olefile
import zlib
import struct

def get_hwp_text(filename):
    f = olefile.OleFileIO(filename)
    dirs = f.listdir()

    # HWP 파일 검증
    if ["FileHeader"] not in dirs or        ["\x05HwpSummaryInformation"] not in dirs:
        raise Exception("Not Valid HWP.")

    # 문서 포맷 압축 여부 확인
    header = f.openstream("FileHeader")
    header_data = header.read()
    is_compressed = (header_data[36] & 1) == 1

    # Body Sections 불러오기
    nums = []
    for d in dirs:
        if d[0] == "BodyText":
            nums.append(int(d[1][len("Section"):]))
    sections = ["BodyText/Section"+str(x) for x in sorted(nums)]

    # 전체 text 추출
    text = ""
    for section in sections:
        bodytext = f.openstream(section)
        data = bodytext.read()
        if is_compressed:
            unpacked_data = zlib.decompress(data, -15)
        else:
            unpacked_data = data
    
        # 각 Section 내 text 추출    
        section_text = ""
        i = 0
        size = len(unpacked_data)
        while i < size:
            header = struct.unpack_from("<I", unpacked_data, i)[0]
            rec_type = header & 0x3ff
            rec_len = (header >> 20) & 0xfff

            if rec_type in [67]:
                rec_data = unpacked_data[i+4:i+4+rec_len]
                section_text += rec_data.decode('utf-16')
                section_text += "\n"

            i += 4 + rec_len

        text += section_text
        text += "\n"

    return text


# In[93]:


decoded_text=get_hwp_text("자소서 참고.hwp")


# In[96]:


import pandas as pd
pd.set_option('display.max_rows', 10000)
items = decoded_text.replace("*", "")
items = decoded_text.replace("\r", "  ")
items = items.split("\n")
list_df = pd.DataFrame(items)
#list_df =  list_df.drop([0], axis = 0)
dfStyler = list_df.style.set_properties(**{'text-align': 'left'})
dfStyler.set_table_styles([dict(selector='th', props=[('text-align', 'left')])])


# In[68]:


df_test = pd.DataFrame(list_df)

df_test.to_clipboard(sep='\t', index=False)

