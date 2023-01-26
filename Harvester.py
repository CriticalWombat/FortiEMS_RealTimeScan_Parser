import zipfile
import os
import patoolib
import xlsxwriter

class grainWrangler:
    def __init__(self):
        self.WB=xlsxwriter.Workbook('GrainReport.xlsx')
        self.WS=self.WB.add_worksheet()
        self.WS.write('A1', 'Hostname')
        self.WS.write('B1', 'Events')

        self.row=0
        self.col=0

        for zip in os.listdir('.'):
            if zip.endswith('.zip'):
                hostname = self.__extractor(zip)
                self.__cleanerUpper(zip)
                events = self.__logWrangler(hostname)
                self.__excelOps(hostname, events)
        self.WB.close()

    def __logWrangler(self, hostname):
        log = f'{hostname}\FCDiagData\general\logs\\realtime_scan.log'
        with open(log, 'r') as rts:
            lines = rts.readlines()
            eventlist = []
            for row in lines:
                key="virus"
                if row.find(key) != -1:
                    eventlist.append(row)
            return eventlist

    def __extractor(self, zip):
        with zipfile.ZipFile(zip, mode="r") as zip:
            zip.extractall()

        for cab in os.listdir('.'):
            if cab.endswith('.cab'):
                hostname=cab.split("_")[1]
                patoolib.extract_archive(cab, outdir=hostname)
                return hostname

    def __cleanerUpper(self, zip):
        for cab in os.listdir('.'):
            if cab.endswith('.cab'):
                os.remove(cab)
#            os.remove(zip)

    def __excelOps(self, hostname, events):
        for virus in (events):
            self.WS.write(self.row, self.col, hostname)
            self.WS.write(self.row, self.col +1, virus)
            self.row +=1

if __name__=='__main__':
    grainWrangler()
