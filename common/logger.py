import logging

LOG_FILE_NAME = 'vinda.log'

logger = logging.getLogger("vinda")
logger.setLevel(logging.DEBUG)

vlog = logging.FileHandler(LOG_FILE_NAME, "a", encoding="utf-8")
vlog.setLevel(logging.DEBUG)
formatter = logging.Formatter('[ %(asctime)s %(filename)s line:%(lineno)d %(levelname)s ]: %(message)s')
vlog.setFormatter(formatter)

logger.addHandler(vlog)