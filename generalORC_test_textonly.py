import hmac
import base64
import hashlib
import json
import time
import uuid
import requests
from urllib.parse import urlparse

def sha256(content):
    x = hashlib.sha256()
    x.update(content.encode())
    return x.hexdigest().upper()

def hmac_sha256(key, content):
    sign = hmac.new(key, content, digestmod="sha256").digest()
    ret = base64.b64encode(sign)
    return ret

# 计算签名
def get_signature(ak, sk, app_key, params):
    request_id = str(uuid.uuid1())
    now_time = time.localtime()
    eop_date = time.strftime("%Y%m%dT%H%M%SZ", now_time)
    eop_date_simple = time.strftime("%Y%m%d", now_time)

    camp_header = "appkey:{0}\nctyun-eop-request-id:{1}\neop-date:{2}\n".format(app_key, request_id, eop_date)

    parsed_url = urlparse(request_url)
    query = parsed_url.query
    query_params = sorted(query.split("&"))
    after_query = ""
    for query_param in query_params:
        if len(after_query) < 1:
            after_query += query_param
        else:
            after_query += "&" + query_param
    content_hash = sha256(json.dumps(params)).lower()

    pre_signature = camp_header + "\n" + after_query + "\n" + content_hash

    k_time = hmac_sha256(sk.encode("utf-8"), eop_date.encode("utf-8"))
    k_ak = hmac_sha256(base64.b64decode(k_time), ak.encode("utf-8"))
    k_date = hmac_sha256(base64.b64decode(k_ak), eop_date_simple.encode("utf-8"))

    # 签名的使用
    signature = hmac_sha256(base64.b64decode(k_date), pre_signature.encode("utf-8"))
    sign_header = "{0} Headers=appkey;ctyun-eop-request-id;eop-date Signature={1}".format(ak, signature.decode())

    # 返回request-id eop-date和sign_header
    return request_id, eop_date, sign_header

# 向服务发送请求
def do_post(url, headers, params):
    response = requests.post(url, data=json.dumps(params), headers=headers)
    try:
        print(response.status_code)
        print(response.json())
        
    except AttributeError:
        print("请求失败")

if __name__ == '__main__':
    request_url = "https://ai-global.ctapi.ctyun.cn/v1/aiop/api/2f3p1pnxpqm8/ocrdetect/ocr/v1/image.json"
    ctyun_ak = 'f6300e68363f48ceadcc88d94581f7c2'
    ctyun_sk = 'cfae02a626d94cccadf5d936765bf010'
    ai_app_key = '3367cd0425f7d818996a69b032032366'

    # 打开图片文件
    f = open(r'test2.png', 'rb')
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