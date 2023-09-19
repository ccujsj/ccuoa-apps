import logging
import logging.handlers
from os import mkdir
from os.path import join,  exists

class Logger(object):
    def __init__(self, logname, backup_count=10):
        self.logname = logname
        self.log_dir = join('.', "logs")
        self.log_file = join(self.log_dir, "{0}.log".format(self.logname))
        self._levels = {
            "DEBUG": logging.DEBUG,
            "INFO": logging.INFO,
            "WARNING": logging.WARNING,
            "ERROR": logging.ERROR,
            "CRITICAL": logging.CRITICAL,
        }
        self._logfmt = "%Y-%m-%d %H:%M:%S"
        self._logger = logging.getLogger(self.logname)
        if not exists(self.log_dir):
            mkdir(self.log_dir)

        LOGFMT = (
                "[ %(levelname)s ] %(threadName)s %(asctime)s "
                "%(filename)s:%(lineno)d %(message)s"
            )
        stream_handler = logging.StreamHandler()
        handler = logging.FileHandler(self.log_file)
        handler.suffix = "%Y%m%d"

        formatter = logging.Formatter(LOGFMT, datefmt=self._logfmt)
        handler.setFormatter(formatter)
        stream_handler.setFormatter(formatter)
        self._logger.addHandler(handler)
        self._logger.addHandler(stream_handler)
        self._logger.setLevel(self._levels.get("DEBUG"))

    @property
    def getLogger(self):
        return self._logger

logger = Logger("Apps").getLogger