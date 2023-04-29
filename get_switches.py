import argparse
import logging

def main():
    logging.basicConfig(level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s', handlers=[logging.FileHandler('automation.log'),logging.StreamHandler()])
    parser = argparse.ArgumentParser(description='Python script to get switches configuration')
    #'Folder with Excel Files'
    parser.add_argument('-e', '--excel-folder', dest='spreadsheets_folder', required=False, default='spreadsheets/', help=argparse.SUPPRESS)
    #'Devices file with Excel Files'
    parser.add_argument('-d', '--devices-file', dest='devices_file', required=False, default='devices.xlsx', help=argparse.SUPPRESS)
    #'nornir configuration file'
    parser.add_argument('-n', '--nornir-config', dest='nornir_config', required=False, default='nornir_config.yaml', help=argparse.SUPPRESS)


if __name__ == '__main__':
    main()