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

class PlaylistClientService (win32serviceutil.ServiceFramework):
	_svc_name_ = "playlist-client"
	_svc_display_name_ = "playlist-client"
	
	def __init__(self, args):
		win32serviceutil.ServiceFramework.__init__(self,args)
		self.isAlive = True
		self.artist = None
		self.title = None
		self.started_time = None
		self.type = None
		self.path = dirname(__file__)
		config = ConfigParser.RawConfigParser()
		config.read(self.path + '\playlistclient.cfg')
		self.interval = float(config.get('Default', 'interval'))
		self.playlist_file = config.get('Default', 'playlist_file')
		# self.loglevel = config.get('Default', 'loglevel')
		logging.basicConfig(filename=self.path + '\playlistclient.log', level=logging.DEBUG)

	def SvcDoRun(self):
		while self.isAlive:
			self.parse_playlist_xml()
			sleep(self.interval)
	
	def SvcStop(self):
		self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
		self.isAlive = False

	def parse_playlist_xml(self):
		id3 = None
		try:
			tree = ET.parse(self.playlist_file)
			root = tree.getroot()
			onair = root.find('OnAir')
			curins = onair.find('CurIns')
			id3 = curins.find('ID3')
		except IOError:
			logging.error('Cannot read playlist file.')
			pass
		except Exception as e:
			logging.exception(e)
			pass
		if id3 is not None:
<<<<<<< HEAD
			if ((id3.get('Artist') != self.artist) or (id3.get('Title') != self.title) or (curins.find('Type').text != self.type)):
				self.artist = id3.get('Artist')
				unicode(self.artist, "utf-8")
				self.title = id3.get('Title')
				unicode(self.title, "utf-8")
				self.started_time = curins.find('StartedTime').text
				self.type = curins.find('Type').text
				self.update_playlist()
=======
			try:
				if ((id3.get('Artist') != self.artist) or (id3.get('Title') != self.title) or (curins.find('Type').text != self.type)):
					self.artist = id3.get('Artist')
					self.title = id3.get('Title')
					self.started_time = curins.find('StartedTime').text
					self.type = curins.find('Type').text
					self.update_playlist()
			except Exception as e:
				logging.exception(e)
				pass
>>>>>>> e43a0757c3a2e0046905fd0253b235f3340229d8

	def update_playlist(self):
		payload = { 'started_time' : self.convert_playlist_time(self.started_time) }

		if self.type == '1':
			if self.title[0:2] != '_n_':
				payload['title'] = self.title
				if self.artist != '':
					payload['artist'] = self.artist
			else:
				payload['title'] = u'Bloco Comercial'
		elif self.type == '0':
			payload['title'] = u'Bloco Comercial'
		elif self.type == '2':
			payload['title'] = u'Ao Vivo!'
		elif self.type == '3':
			payload['title'] = u'Ao Vivo!'
		elif self.type == '4':
			payload['title'] = u'Hora Certa!'
		else:
			payload['title'] = u'Sem Informação'
		try:
			res = requests.post('http://playlist-service.appspot.com/v1/playlist/add', data=payload)
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