import hmac
import base64
import hashlib
import json
import time
import uuid
import requests
import pandas as pd
from urllib.parse import urlparse

def sha256(content):
    x = hashlib.sha256()
    x.update(content.encode())
    return x.hexdigest().upper()

def hmac_sha256(key, content):
    sign = hmac.new(key, content, digestmod="sha256").digest()
    ret = base64.b64encode(sign)
    return ret

# 保存为显示格式
def save_json_to_csv(response_json, filename='response_data.csv'):
    """
    将 JSON 数据整理成表格并保存为 CSV 文件。

    :param response_json: 包含 OCR 结果的 JSON 数据
    :param filename: 保存的 CSV 文件名
    """
    records = []
    for obj in response_json.get('returnObj', []):
        text_line = obj.get('text_line', 0)
        for detail in obj.get('detail', []):
            record = {
                'name': detail.get('name', ''),
                'text': detail.get('text', ''),
                'box': ','.join(map(str, detail.get('box', []))),
                'text_line': text_line
            }
            records.append(record)

    # 创建 DataFrame
    df = pd.DataFrame(records)

    # 保存为 CSV 文件
    df.to_csv(filename, index=False, encoding='utf-8-sig')
    print(f"表格已保存为 {filename}")

# 计算签名

def get_signature(ak, sk, app_key, params):
    # 创建待签名字符串
    # 一、header部分
    # 主要包括3个header需要作为签名内容：appkey、ctyun-eop-request-id、eop-date
    # 1. 首先通过uuid生成ctyun-eop-request-id
    request_id = str(uuid.uuid1())

    # 2. 获取当前时间戳并对时间进行格式化
    now_time = time.localtime()
    eop_date = time.strftime("%Y%m%dT%H%M%SZ", now_time)
    eop_date_simple = time.strftime("%Y%m%d", now_time)

    # 3. 对header部分按照字母顺序进行排序并格式化
    camp_header = "appkey:{0}\nctyun-eop-request-id:{1}\neop-date:{2}\n".format(app_key, request_id, eop_date)

    # 二、query部分
    # 对url的query部分进行排序
    parsed_url = urlparse(request_url)
    query = parsed_url.query
    query_params = sorted(query.split("&"))
    after_query = ""

    for query_param in query_params:
        if len(after_query) < 1:
            after_query += query_param
        else:
            after_query += "&" + query_param

    # 三、body参数进行sha256摘要
    # sha256 body
    content_hash = sha256(json.dumps(params)).lower()
    # 完成创建待签名字符串
    pre_signature = camp_header + "\n" + after_query + "\n" + content_hash
    # 构造动态密钥
    k_time = hmac_sha256(sk.encode("utf-8"), eop_date.encode("utf-8"))
    k_ak = hmac_sha256(base64.b64decode(k_time), ak.encode("utf-8"))
    k_date = hmac_sha256(base64.b64decode(k_ak), eop_date_simple.encode("utf-8"))

    # 签名的使用
    signature = hmac_sha256(base64.b64decode(k_date), pre_signature.encode("utf-8"))

    # 将数据整合得到真正的header中的内容
    sign_header = "{0} Headers=appkey;ctyun-eop-request-id;eop-date Signature={1}".format(ak, signature.decode())

    # 返回request-id eop-date和sign_header
    return request_id, eop_date, sign_header

# 向服务发送请求

def do_post(url, headers, params):
    response = requests.post(url, data=json.dumps(params), headers=headers)
    try:
        print(response.status_code)
        print(response.json()) 
        with open('response.json', 'w') as f:
            json.dump(response.json(), f, ensure_ascii=False, indent=4)
        # save_json_to_csv(response.json())
    except AttributeError:
        print("请求失败")

if __name__ == '__main__':
    request_url = "https://ai-global.ctapi.ctyun.cn/v1/aiop/api/2f3p1pnxpqm8/ocrdetect/ocr/v1/image.json"
    ctyun_ak = 'f6300e68363f48ceadcc88d94581f7c2'
    ctyun_sk = 'cfae02a626d94cccadf5d936765bf010'
    ai_app_key = '3367cd0425f7d818996a69b032032366'

    # 打开图片文件
    f = open(r'test4.png', 'rb')
    img_base64 = base64.b64encode(f.read()).decode()

    # body内容
    params = {"imageContent": img_base64}
    params = {"data": [img_base64] }
    # 调用get_signature方法获取签名

    request_id, eop_date, sign_header = get_signature(ctyun_ak, ctyun_sk, ai_app_key, params)
    
    headers = {
        'Content-Type': 'application/json;charset=UTF-8',
        'ctyun-eop-request-id': request_id,
        'appkey': ai_app_key,
        'Eop-Authorization': sign_header,
        'eop-date': eop_date,
        'host': 'ai-global.ctapi.ctyun.cn'
    }

    print("请求头部:")
    print(headers)

    # 执行post请求
    do_post(request_url, headers, params)

