import cv2

# 元画像読み込み
img=cv2.imread('./samplePics/space_1.jpg',0)
# 元画像を表示
cv2.imshow("original",img)

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

print('White_Area='+str(white_area/whole_area*100)+'%')
print('Black_Area='+str(black_area/whole_area*100)+'%')
print('Gray_Area='+str(gray_area/whole_area*100)+'%')        
            
# 3値化した画像の表示
cv2.imshow("create",img)
cv2.waitKey(0)
cv2.destroyAllWindows()