FileNameList = ["tle.txt"]


for FileName in FileNameList :
    print(FileName)

    TleFile = open(FileName, "r")
    TLE = TleFile.read()
    TleFile.close()

    TLE = TLE.replace(
    """NEPALISAT1
1 44329U """,
    """RAAVANA1
1 44329U """)

    TLE = TLE.replace(
    """RAAVANA1
1 44330U """,
    """UGUISU
1 44330U """)

    TLE = TLE.replace(
    """UGUISU
1 44331U """,
    """NEPALISAT1
1 44331U """)

    print(TLE)

    TleFile = open(FileName, 'w')
    TleFile.write(TLE)
    TleFile.close()

#    print('TLE rename Success\n')


print("All TLE rename Success!")
