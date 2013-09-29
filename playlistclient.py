import win32serviceutil
import win32service
import logging
from time import sleep, strptime, mktime
from os.path import dirname
import ConfigParser
import requests
import xml.etree.ElementTree as ET
from calendar import timegm
from datetime import datetime
import pytz


class Item:
    def __init__(self, artist, title, started_time, type):
        self.artist = artist
        self.title = title
        self.started_time = started_time
        self.type = type
        if self.type == '1':
            if self.title[0:3] == '_n_':
                self.artist = None
                self.title = u'Bloco Comercial'
        elif self.type == '0':
            self.artist = None
            self.title = u'Bloco Comercial'
        elif self.type == '2':
            self.artist = None
            self.title = u'Ao Vivo!'
        elif self.type == '3':
            self.artist = None
            self.title = u'Ao Vivo!'
        elif self.type == '4':
            self.artist = None
            self.title = u'Hora Certa!'
        else:
            self.artist = None
            self.title = u'Sem Informa\xe7\xe3o'

    def __cmp__(self, other):
        if other is None:
            return -1
        if self.type != other.type:
            return -1
        if (self.artist != other.artist) or (self.title != other.title):
            return -1
        return 0


class PlaylistClientService(win32serviceutil.ServiceFramework):
    _svc_name_ = "playlist-client"
    _svc_display_name_ = "playlist-client"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.is_alive = True
        self.last_item = None
        self.path = dirname(__file__)
        config = ConfigParser.RawConfigParser()
        config.read(self.path + '\playlistclient.cfg')
        self.interval = float(config.get('Default', 'interval'))
        self.playlist_file = config.get('Default', 'playlist_file')
        logging.basicConfig(filename=self.path + '\playlistclient.log', level=logging.DEBUG)

    def SvcDoRun(self):
        while self.is_alive:
            self.parse_playlist_xml()
            sleep(self.interval)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.is_alive = False

    def parse_playlist_xml(self):
        id3 = None
        try:
            tree = ET.parse(self.playlist_file)
            root = tree.getroot()
            onair = root.find('OnAir')
            curins = onair.find('CurIns')
            started_time = curins.find('StartedTime').text
            type = curins.find('Type').text
            id3 = curins.find('ID3')
        except IOError:
            logging.error('Cannot read playlist file.')
            pass
        except Exception as e:
            logging.exception(e)
            pass
        if id3 is not None:
            try:
                artist = id3.get('Artist')
                title = id3.get('Title')
                item = Item(artist, title, started_time, type)
                self.update_playlist(item)
            except Exception as e:
                logging.exception(e)
                pass

    def update_playlist(self, item):
        if item != self.last_item:
            self.last_item = item

            payload = {
                'started_time': self.convert_playlist_time(item.started_time),
                'artist': item.artist,
                'title': item.title
            }

            try:
                requests.post('http://playlist-service.appspot.com/v1/playlist/add', data=payload)
                logging.debug('Updating Playlist:')
                logging.debug(str(payload))
            except Exception as e:
                logging.exception(e)
                pass

    def convert_playlist_time(self, str_time):
        timestamp = None
        try:
            time = strptime(str_time, '%d/%m/%Y %H:%M:%S')
            brt = pytz.timezone('America/Sao_Paulo')
            dt = datetime.fromtimestamp(mktime(time))
            brt_dt = brt.localize(dt)
            utc_dt = brt_dt - brt_dt.utcoffset()
            utc_dt = utc_dt.replace(tzinfo=pytz.utc)
            timestamp = timegm(utc_dt.utctimetuple())
        except Exception as e:
            logging.exception(e)
            pass
        return timestamp


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(PlaylistClientService)