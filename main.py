# -*- coding: utf-8 -*-

import socket
import struct
import win32print
import yaml
import logging
from logging import handlers

__author__ = 'Krzysztof Karolak'
__copyright__ = '© Copyright 2020'
__version__ = '0.1'

"""
Python 2.7 required

- Load printers and locations from YAML configuration file.
- Detect local network
- Load local printers and compare with known locations by network addresses
- Set default printer for detected location (if printer is not installed, do nothing)

"""


def config_from_file(fileyml):
    # Load ip's, ports and locations to multi-dimension array or list
    with open(fileyml, 'r') as stream:
        try:
            config_list = yaml.safe_load(stream)
            return config_list
        except yaml.YAMLError as exc:
            print(exc)
            return False


# Network compare methods
def make_mask(n):
    return (2L << n-1) - 1


def dotted_quad_to_num(ip):
    return struct.unpack('L', socket.inet_aton(ip))[0]


def network_mask(ip, bits):
    return dotted_quad_to_num(ip) & make_mask(bits)


def address_in_network(ip, net):
    return ip & net == net


def detect_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.connect(("8.8.8.8", 80))
    local_ip = s.getsockname()[0]
    s.close()
    logger.info("Detected local ip: " + local_ip)
    return local_ip


def detect_current_location(prt_ports):
    targetPrinter = None
    local_ip = dotted_quad_to_num(detect_local_ip())
    for printer_port in prt_ports:
        if address_in_network(local_ip, network_mask(prt_ports[printer_port]['networkAddress'], prt_ports[printer_port]['networkMask'])):
            targetPrinter = prt_ports[printer_port]['printerPort']
            logger.info("Detected location " + prt_ports[printer_port]['networkLocation'] + ", selecting port " + targetPrinter)
    return targetPrinter


def set_default_printer(printer_port):
    targetPrinterName = None
    local_printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS, None, 5)

    logger.info(local_printers)

    for lp in local_printers:
        if printer_port == lp['pPortName']:
            targetPrinterName = lp['pPrinterName']
            logger.info("Set default printer to " + lp['pPrinterName'])

    # Set default printer by name
    if targetPrinterName:
        win32print.SetDefaultPrinter(targetPrinterName)
    else:
        logger.error("Could not find default printer.")


config_file = "config.yml"
printer_ports = config_from_file(config_file)

# Data Logging to file
logPath = printer_ports['log_path']
logger = logging.getLogger('my_app')
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

logHandler = handlers.RotatingFileHandler(logPath, maxBytes=50000, backupCount=2)
logHandler.setLevel(logging.DEBUG)
logHandler.setFormatter(formatter)
logger.addHandler(logHandler)

# Set Default Printer
targetPrinterPort = detect_current_location(printer_ports['printers'])
set_default_printer(targetPrinterPort)
