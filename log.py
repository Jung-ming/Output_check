import logging, sys, os


def get_logger(rootName="__main__", childName="", log_dir="logs"):
    logName = rootName if not childName else rootName + "." + childName
    # print("Your Log name is : ", logName)
    logger = logging.getLogger(logName)
    logger.setLevel(logging.DEBUG)

    from datetime import datetime
    now = datetime.now()
    dt_string = now.strftime("%Y%m%d")
    fileName = dt_string

    if not os.path.isdir(log_dir):
        os.mkdir(log_dir)

    fileName = os.path.join(log_dir, fileName)

    if not childName:  # 只能在最頂層加，如果每一層都這樣加，每一個 child logger 也會都 print 一行
        # file handler
        fh = logging.FileHandler(fileName + ".log", encoding='utf-8-sig')
        fh.setLevel(logging.INFO)
        ch = logging.StreamHandler()  # sys.stdout
        ch.setLevel(logging.DEBUG)

        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s : %(message)s')
        # formatter = logging.Formatter('%(message)s')
        ch.setFormatter(formatter)
        fh.setFormatter(formatter)

        # put filehandler into logger
        logger.addHandler(fh)
        logger.addHandler(ch)

    return logger
