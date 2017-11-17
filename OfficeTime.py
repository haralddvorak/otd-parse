import datetime
import csv

MAC_OS_EPOCH = -2082823200  # OfficeTime/REALbasic uses the classic Mac OS's 1/1/1904 epoch for timestamps


class OfficeTimeFile:
    def __init__(self, path):
        self.path = path
        self.all_projects = []
        self.all_sessions = []

        f = open(path, 'r')
        contents = f.read()
        f.close()

        major_sections = contents.split('###########')
        header = major_sections[0]

        # project_explanation = major_sections[1]
        # session_explanation = major_sections[2]
        actual_data = major_sections[3]
        # device_sync_info = major_sections[4]
        # deleted_items_explanation = major_sections[5]
        # deleted_items = major_sections[7]

        data_objects = actual_data.split('########')

        # Parse the file line by line, making projects and sessions as we encounter them
        last_project = None
        num = 0
        for obj_text in data_objects:
            # print(obj_text)
            obj_lines = obj_text.splitlines()
            first_line = obj_lines[1]
            # print(first_line)
            if first_line == 'SESSION':  # This is a session belonging to the last project
                # print("session found")
                cat_name = obj_lines[8]
                ticked = self._parse_bool(obj_lines[9])

                s = OfficeTimeFile.Session(uid=obj_lines[21],
                                           project=last_project,
                                           start_time=self._parse_timestamp(obj_lines[2]),
                                           end_time=self._parse_timestamp(obj_lines[20]),
                                           length=datetime.timedelta(seconds=float(obj_lines[3])),
                                           adjustment=datetime.timedelta(float(obj_lines[4])),
                                           notes=obj_lines[5])

                self.all_sessions.append(s)
                last_project.sessions.append(s)

            elif first_line.startswith('###Project'):  # this is a project definition
                created = obj_lines[6]
                modified = obj_lines[21]

                p = OfficeTimeFile.Project(uid=obj_lines[26],
                                           name=obj_lines[2],
                                           client=obj_lines[4])
                p.archived = self._parse_bool(obj_lines[28])

                self.all_projects.append(p)
                last_project = p

            else:  # this is something else, skip it
                pass

            num += 1

    def _parse_bool(self, string):
        return string == 'True'

    def _parse_timestamp(self, string):
        return datetime.datetime.fromtimestamp(float(string) + MAC_OS_EPOCH)

    def csv_export(self):
        path = self.path
        print(path)
        export_filename = path.replace('.otd', '.csv')
        print(export_filename)
        # Create new file/overwrite existing file

        with open(export_filename, 'w', newline='', encoding='utf-8') as csv_file:
            csv_file.write('\ufeff')  # BOM character for Excel to recognize file format utf8

        with open(export_filename, 'a', newline='', encoding='utf-8') as csv_file:
            csv_writer = csv.writer(csv_file, dialect='excel', delimiter=';', skipinitialspace=True)
            for p in self.all_projects:
                line_out = ([p.name, '', '', ''])
                csv_writer.writerow(line_out)

                line_out = (['Start', 'Ende', 'Dauer', 'Notiz'])
                csv_writer.writerow(line_out)

                for s in p.sessions:
                    start_time = s.start_time.strftime('%d.%m.%Y %H:%M:%S')
                    end_time = s.end_time.strftime('%d.%m.%Y %H:%M:%S')
                    duration = s.length
                    notes = s.notes.replace("[EMPTY_LINE]", "Keine Notiz")
                    line_out = ([start_time, end_time, duration, notes])
                    csv_writer.writerow(line_out)

    class Project:
        def __init__(self, uid, name, client):
            self.uid = uid
            self.name = name
            self.client = client
            self.archived = False
            self.sessions = []

        def __str__(self):
            return "<Project: %s, %s>" % (self.name, self.uid)

    class Session:
        def __init__(self, uid, project, start_time, end_time, length, adjustment=datetime.timedelta(0), notes=''):
            self.uid = uid
            self.project = project
            self.start_time, self.end_time, self.length, self.adjustment = start_time, end_time, length, adjustment
            self.notes = notes

        def __str__(self):
            return '<Session: {0}: {1} -> {2} für {3}>'.format(self.project.name, self.notes, self.start_time,
                                                               self.length)
