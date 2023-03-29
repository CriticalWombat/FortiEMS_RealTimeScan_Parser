import zipfile
import os
import patoolib
import xlsxwriter
from termcolor import colored


class FortiEMS_AV_Parser:
    def __init__(self):
        self.WB=xlsxwriter.Workbook('AV_Events.xlsx')
        self.WS=self.WB.add_worksheet()
        self.WS.write('A1', 'Hostname')
        self.WS.write('B1', 'Events')
        self.row=0
        self.col=0

        for file in os.listdir('.'):
            if file.endswith('.zip'):
                self.__zipExtractor(file)
            else:
                continue

        for file in os.listdir('.'):
            if file.endswith('.cab'):
                hostname = file.split("_")[1]
                self.__cabExtractor(file, hostname)
                events = self.__logWrangler(hostname)
                self.__excelOps(hostname, events)
            else:
                continue
        self.__removeExtracted()
        self.WB.close()
            


    def __logWrangler(self, hostname):
        log = f'{hostname}\FCDiagData\general\logs\\realtime_scan.log'
        if os.path.isfile(log):
            with open(log, 'r' encoding="UTF-8") as rts:
                lines = rts.readlines()
                eventlist = []
                for row in lines:
                    key="virus"
                    if row.find(key) != -1:
                        eventlist.append(row)
                return eventlist
        else:
            error = colored(f'No Real Time Scan log found for {hostname}...', 'red')
            print(error)
            

    def __zipExtractor(self, zip):
        with zipfile.ZipFile(zip, mode="r") as zip:
            return zip.extractall()

    def __cabExtractor(self, file, hostname):
            try:
                patoolib.extract_archive(file, outdir=hostname, verbosity=-1)
                return
            except:
                error = colored(f'Issue with extracting {hostname}, please review folder manually!', 'red')    
                print(error)
                os.rename(hostname, "ERROR-{hostname}".format(hostname=hostname))
                return
            
            
    def __removeExtracted(self):
        for cab in os.listdir('.'):
            if cab.endswith('.cab'):
                os.remove(cab)

    def __excelOps(self, hostname, events):
        if events:
            for virus in (events):
                self.WS.write(self.row, self.col, hostname)
                self.WS.write(self.row, self.col +1, virus)
                self.row +=1

if __name__=='__main__':
    FortiEMS_AV_Parser()
