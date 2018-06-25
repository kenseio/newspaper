#!/usr/bin/env python
# -*- coding: utf-8 -*-

import requests
import json

with open('secret.json') as sf:
    data = json.load(sf)

token = data['line_token']
url = 'https://notify-api.line.me/api/notify'
# header = {'Content-Type':'multipart/form-data', 'Authorization':'Bearer ' + token}
header = {'Authorization': 'Bearer ' + token}
option = {'message': '\n新聞配信完了しました！\nKindleでダウンロードしてください', 'stickerPackageId': 1, 'stickerId': 106}

response = requests.post(url, data=option, headers=header)
# print(json.load(response))
print(response.text)
