#!/usr/bin/env python
# -*- coding: utf-8 -*-

import time
from PIL import Image
from robobrowser import RoboBrowser


# 画像の下処理。サイズを小さくして、グレースケールにする。クオリティは変えない
def image_process(root, img_src):
    br = RoboBrowser(parser='lxml', user_agent='a python robot2', history=True)
    img_path = root + 'img.jpg'
    request = br.session.get(img_src, stream=True)

    try:
        with open(img_path, "wb") as img_file:
            img_file.write(request.content)

        time.sleep(0.3)
        im = Image.open(img_path)
        # print(im.format, im.size, im.mode)

        if im.size[0] > 500:
            rate = 500 / im.size[0]
        else:
            rate = 1
        # print(str(rate))
        size = (int(im.size[0] * rate), int(im.size[1] * rate))
        new_im = im.resize(size).convert('L')
        # print(new_im.format, new_im.size, new_im.mode)
        new_im.save(img_path)

    except:
        print('/---Warning:画像読み込みできなかったよ')


if __name__ == '__main__':
    print("このコードはインポートして使ってね。")
