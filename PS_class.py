import datetime
from collections import namedtuple
from dataclasses import dataclass


@dataclass
class PS:
    pk: int
    Surveyor: str
    Zone: str
    Type: str
    Street: str
    Section: str
    Side: str
    Restriction: str
    LatLong: tuple
    Plates: list
    Count: list
    UniqueCount: list
    Times: list
    Dates: list

    def __str__(self):
        return 'PK: {}\nSurveyor: {}\nZone: {}\nType: {}\nStreet: {}\
                \nSection: {}\nSide: {}\nRestriction: {}\nLatLong: {}\
                \nDates: {}\nTimes: {}\nPlates: {}\nCount: {}\nUnique Count: {}'\
                .format(str(self.pk), self.Surveyor, self.Zone, self.Type,\
                        self.Street, self.Section, self.Side, self.Restriction,\
                        str(self.LatLong), str(self.Dates), str(self.Times),\
                        str(self.Plates), str(self.Count), str(self.UniqueCount))
