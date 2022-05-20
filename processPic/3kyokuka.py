from importlib.metadata import files
from tkinter import filedialog
import cv2

class ProcessingPic:
    def __init__(self, fle):
        # 元画像読み込み
        img=cv2.imread(fle,0)
        # 元画像を表示
        # cv2.imshow("original",img)

        color_area = self.threeColor(img)

        self.black_area = color_area['black_area']
        self.white_area = color_area['white_area']
        self.gray_area = color_area['gray_area']
    
    def threeColor(self, img):
        # 画像の縦の画素数、横の画素数を取得
        height,width=img.shape

        # 黒にする上限を指定
        th_low_value=25
        # 白にする上限を指定
        th_high_value=150
        # 灰色を定義
        gray=100
        # 元画像全体の画素数
        whole_area=img.size

        # 三色の面積初期化
        white_area=0
        black_area=0
        gray_area=0


        # 画像の３値化と各色の面積比を求める
        for i in range(height):
            for j in range(width):
                if img[i,j]<th_low_value:
                    img[i,j]=0
                    black_area+=1
                elif th_low_value <= img[i,j] <=th_high_value:
                    img[i,j]=gray
                    gray_area+=1
                else :
                    img[i,j]=255
                    white_area+=1

        # 3値化した画像の表示
        # cv2.imshow("create",img)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()

        dic = {
            'black_area':black_area/whole_area*100, 
            'white_area':white_area/whole_area*100, 
            'gray_area':gray_area/whole_area*100
        }

        return dic

def selectPic():
    typ = [('画像ファイル','*.jpeg;*.jpg;*.png')] 
    dir = '.\samplePics'
    fles = filedialog.askopenfilenames(filetypes = typ, initialdir = dir)
    return fles

def main():
    fles = selectPic()
    pics = {}
    for fle in fles:
        dic_name = fle.rsplit('/', 1)[1]
        pics[dic_name] = ProcessingPic(fle)
        print(dic_name, "|",round(pics[dic_name].black_area, 2))

        # if pics[dic_name].black_area > 80:
        #     print("ひっくり返ってます")
        # else:
        #     print("正常です")

if __name__ == '__main__':
    main()