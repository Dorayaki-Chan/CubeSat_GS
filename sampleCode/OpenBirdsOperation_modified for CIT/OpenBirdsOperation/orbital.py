from datetime import datetime

from pyorbital.orbital import Orbital
from pyorbital.tlefile import Tle


# TLEの読み込み
aqua_tle = Tle('AQUA', 'aqua.tle')
aqua_orbit = Orbital('AQUA', line1=aqua_tle.line1, line2=aqua_tle.line2)

now = datetime.utcnow()

# 現在の緯度、経度、高度を取得
lon, lat, alt = aqua_orbit.get_lonlatalt(now)
print('Aquaの現在地')
print('経度: ', lon)
print('緯度: ', lat)
print('高度[km]: ', alt)
print('')

# 24時間以内に東京タワーから衛星が見える時間を計算
pass_time_list = (aqua_orbit.get_next_passes(utc_time=now, length=24,
                                             lon=139.75, lat=35.66, alt=0.333))
print('次にAquaが到来する時刻[UTC]: ',
      pass_time_list[0][0].strftime('%Y/%m/%d %H:%M:%S'))
