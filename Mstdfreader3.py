"""
The MIT License (MIT)
Copyright (c) 2016 Cahyo Primawidodo

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
documentation files (the "Software"), to deal in the Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and
to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of
the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO
THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

__author__ = 'cahyo primawidodo 2016'
Revised by cdhsu2001@gmail.com in 2020
"""

import time
import json  
import io
import struct 
import logging 
import re
import math  

class Reader:
    HEADER_SIZE = 4 

    def __init__(self, stdf_ver_json=None):
        self.logger = logging.getLogger(self.__class__.__name__)
        self.logger.setLevel(logging.DEBUG)
        format = '%(levelname)s-%(name)s: %(message)s'
        formatter = logging.Formatter(format)   
        
        logfile = './Mstdflogger3.log'
        filehandler = logging.FileHandler(logfile, mode='w')
        filehandler.setFormatter(formatter)
        self.logger.addHandler(filehandler)
        
        self.STDF_TYPE = {} 
        self.STDF_IO = io.BytesIO(b'') 
        self.REC_NAME = {} 
        self.FMT_MAP = {} 
        self.e = '<' 
        self.body_start = 0 
        self.load_byte_fmt_mapping() 
        self.load_stdf_type(json_file=stdf_ver_json)

    def load_stdf_type(self, json_file):

        if json_file is None:
            input_file = './stdf_ver4.json' 
        self.logger.info('Loading STDF configuration file > {}'.format(input_file)) 
        
        with open(input_file) as fp:
            self.STDF_TYPE = json.load(fp) 

        for k, v in self.STDF_TYPE.items(): 
            typ_sub = (v['rec_typ'], v['rec_sub'])
            self.REC_NAME[typ_sub] = k

    def load_byte_fmt_mapping(self):
        self.FMT_MAP = {
            "U1": "B", 
            "U2": "H",
            "U4": "I",
            "U8": "Q",
            "I1": "b",
            "I2": "h",
            "I4": "i",
            "I8": "q",
            "R4": "f",
            "R8": "d",
            "B0": "0s", 
            "B1": "B",
            "C1": "c"
#            "N1": "B"
            }

#======================================================================================================#

    def load_stdf_file(self, stdf_file):
        self.logger.info('Opening STDF file > {}'.format(stdf_file))
        with open(stdf_file, mode='rb') as fs: 
            self.STDF_IO = io.BytesIO(fs.read()) 
        self.logger.info('Detecting STDF file size > {}'.format( len( self.STDF_IO.getvalue() ) ))    

#======================================================================================================#

    def read_and_unpack_header(self):
        header_raw = self.STDF_IO.read(self.HEADER_SIZE) 
        header = False
        if header_raw:
            header = struct.unpack(self.e + 'HBB', header_raw) 
            rec_name = self.REC_NAME.get((header[1], header[2]), 'oops') 
#            self.logger.debug('header={}, rec_name={}'.format(header, rec_name))
            
            if (rec_name == 'FAR') and (header[0] > 2): 
                header = struct.unpack('>' + 'HBB', header_raw)  
#                self.logger.debug('Revised header={}, rec_name={}'.format(header, rec_name))
            
        return header
    
    def read_body(self, rec_size):
        self.body_start = self.STDF_IO.tell()
        body_raw = io.BytesIO( self.STDF_IO.read(rec_size) ) 
        return body_raw 

    def unpack_body(self, header, body_raw):
        rec_len, rec_typ, rec_sub = header 
        typ_sub = (rec_typ, rec_sub) 
        rec_name = self.REC_NAME.get(typ_sub, 'oops')
        max_tell = rec_len
        odd_nibble = True
        body = {} 
        
        if rec_name in self.STDF_TYPE: 
            for field, fmt_raw in self.STDF_TYPE[rec_name]['boddy']:
                if body_raw.tell() >= max_tell: 
                    break

                array_data = []
                if fmt_raw.startswith('K'): 
                    mo = re.match('^K([0xn])(\w{2})', fmt_raw)
                    fmt_act = mo.group(2)                     
                    
                    if field == 'SITE_NUM':
                        n = body['SITE_CNT']  
                    elif field == 'RTST_BIN':
                        n = body['NUM_BINS']  
                    elif field == 'PMR_INDX':
                        n = body['INDX_CNT'] 
                    elif field in ['GRP_INDX', 'GRP_MODE', 'GRP_RADX', 'PGM_CHAR', 'RTN_CHAR', 'PGM_CHAL', 'RTN_CHAL']:
                        n = body['GRP_CNT']   
                    elif field in ['RTN_INDX', 'RTN_STAT']: 
                        n = body['RTN_ICNT']  
                    elif field in ['PGM_INDX', 'PGM_STAT']:
                        n = body['PGM_ICNT']  
                    elif field in ['RTN_STAT', 'RTN_INDX']: 
                        n = body['RTN_ICNT']  
                    elif field in ['RTN_RSLT']:
                        n = body['RSLT_CNT']  
                    
                    for i in range(n):
                        data, odd_nibble = self.get_data(fmt_act, body_raw, odd_nibble)
                        array_data.append(data)

                    body[field] = array_data
                    odd_nibble = True

                elif fmt_raw.startswith('V'): 
                    vn_map = ['B0', 'U1', 'U2', 'U4', 'I1', 'I2',
                              'I4', 'R4', 'R8', 'B0', 'Cn', 'Bn', 'Dn', 'N1']
                    n, = struct.unpack(self.e + 'H', body_raw.read(2))

                    for i in range(n):
                        idx, = struct.unpack(self.e + 'B', body_raw.read(1))
                        fmt_vn = vn_map[idx]

                        data, odd_nibble = self.get_data(fmt_vn, body_raw, odd_nibble)
                        array_data.append(data)

                    body[field] = array_data
                    odd_nibble = True

                else:
                    body[field], odd_nibble = self.get_data(fmt_raw, body_raw, odd_nibble) 

        else:
            self.loggger.warning('Record name={} is not in self.STDF_TYPE'.format(rec_name))

        body_raw.close()
        return rec_name, body

#======================================================================================================# 

    def read_record(self):
        header = self.read_and_unpack_header()

        if header:
            rec_size = header[0]
#            self.logger.debug('BODY start at ={}'.format(self.STDF_IO.tell() ) )
            body_raw = self.read_body(rec_size) 
#            self.logger.debug('BODY end at ={}'.format(self.STDF_IO.tell() ) )
            rec_name, body = self.unpack_body(header, body_raw)

            if rec_name == 'FAR':
                self.set_endian(body['CPU_TYPE'])

            return rec_name, header, body

        else:
            self.logger.info('Closing STDF_IO at ={}'.format(self.STDF_IO.tell() ) )
            self.STDF_IO.close()
            return False

#======================================================================================================#  

    def get_data(self, fmt_raw, body_raw, odd_nibble):
        data = 0
        
        if fmt_raw == 'N1': 
            if odd_nibble:
                nibble, = struct.unpack(self.e + 'B', body_raw.read(1)) 
                _, data = nibble >> 4, nibble & 0xF 
                odd_nibble = False
            else:
                body_raw.seek(-1, 1)
                nibble, = struct.unpack(self.e + 'B', body_raw.read(1)) 
                data, _ = nibble >> 4, nibble & 0xF 
                odd_nibble = True
        else:
            fmt, buf = self.get_format_and_buffer(fmt_raw, body_raw)

            if fmt:
                if fmt.endswith('s') or fmt=='c' :
                    data = str(buf, encoding='utf-8')
                else:
                    d = struct.unpack(self.e + fmt, buf) 
                    data = d[0] if len(d) == 1 else d #!!
            
            odd_nibble = True
            
        return data, odd_nibble

    def get_format_and_buffer(self, fmt_raw, body_raw):
        fmt = self.get_format(fmt_raw, body_raw)
        
        if fmt:
            size = struct.calcsize(fmt)
            buf = body_raw.read(size) 
            return fmt, buf
        else:
            return 0, 0

    def get_format(self, fmt_raw, body_raw):

        if fmt_raw in self.FMT_MAP: 
            return self.FMT_MAP[fmt_raw]

        elif fmt_raw == 'Sn': 
            buf = body_raw.read(2)
            n, = struct.unpack(self.e + 'H', buf)
            posfix = 's'

        elif fmt_raw == 'Cn':
            buf = body_raw.read(1) 
            n, = struct.unpack(self.e + 'B', buf)
            posfix = 's'

        elif fmt_raw == 'Bn':
            buf = body_raw.read(1) 
            n, = struct.unpack(self.e + 'B', buf)
            posfix = 'B'

        elif fmt_raw == 'Dn':
            buf = body_raw.read(2)
            h, = struct.unpack(self.e + 'H', buf)
            n = math.ceil(h/8) 
            posfix = 'B'
        else:
            raise ValueError('To check fmt_raw and body_raw')

        return str(n) + posfix  

#======================================================================================================#  

    def set_endian(self, cpu_type):
        if cpu_type == 1:
            self.e = '>'
        elif cpu_type == 2:
            self.e = '<'
        else:
            self.logger.critical('Value of FAR: CPU_TYPE is not 1 or 2, invalid endian.')
            raise IOError('To check cpu_type')

    @staticmethod 
    def get_multiplier(field, body):
        if field == 'SITE_NUM':
            return body['SITE_CNT']  # SDR (1, 80)

        elif field == 'RTST_BIN':
            return body['NUM_BINS']  # RDR (1, 70)

        elif field == 'PMR_INDX':
            return body['INDX_CNT']  # PGR (1, 62)

        elif field in ['GRP_INDX', 'GRP_MODE', 'GRP_RADX', 'PGM_CHAR', 'RTN_CHAR', 'PGM_CHAL', 'RTN_CHAL']:
            return body['GRP_CNT']  # PLR (1, 63)

        elif field in ['RTN_INDX', 'RTN_STAT']:
            return body['RTN_ICNT']  # FTR (15, 20)

        elif field in ['PGM_INDX', 'PGM_STAT']:
            return body['PGM_ICNT']  # FTR (15, 20)

        elif field in ['RTN_STAT', 'RTN_INDX']:
            return body['RTN_ICNT']  # MPR (15, 15)

        elif field in ['RTN_RSLT']:
            return body['RSLT_CNT']  # MPR (15, 15)

        else:
            raise ValueError

#======================================================================================================#     

    def __iter__(self):
        return self

    def __next__(self):
        repeatobj = self.read_record()
        
        if repeatobj:
            return repeatobj
        else:
            raise StopIteration

#======================================================================================================#    

timerstart = time.time()
r1 = Reader()

STDFfilename = input('Input STDF here, its filename (not zipped) should be file.stdf: ')
r1.load_stdf_file(STDFfilename)
filename = re.search('(.*).stdf', STDFfilename)
XMLfilename = filename.group(1) + '.xml'
fwt = open(XMLfilename, mode='w', encoding='utf-8')

for rec_name, header, body in r1: 
    fwt.writelines(['\n', rec_name, ':'])
    for fieldname, fmtname in r1.STDF_TYPE[rec_name]['boddy']:
        if fieldname in body: 
            if fmtname == ('Cn' or 'C1'):
                fwt.writelines([' ', fieldname, '=', body[fieldname] ])
            else:
                fwt.writelines([' ', fieldname, '=', str(body[fieldname]) ])
       
fwt.close()     
timerend = time.time()
print("It spent {}sec".format(timerend-timerstart) )    
  

