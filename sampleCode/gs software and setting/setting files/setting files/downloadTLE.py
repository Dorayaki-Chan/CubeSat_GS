from spacetrack import SpaceTrackClient

#st = SpaceTrackClient('o108074d@mail.kyutech.jp', 'AobaVeloxIII20170116')
st = SpaceTrackClient(EnvEmail, EnvPassword)
TLE = st.tle_latest(favorites='Kyutech BIRDS GS', ordinal=1, epoch='>now-30', format='3le')
TleFile = open('tle.txt', 'w')
TleFile.write(TLE)
TleFile.close()

 
print('TLE Update Success')

# 参考リンク
### http://rocket-live.blogspot.com/2018/05/orbitrontlespacetrackwebpython-python.html