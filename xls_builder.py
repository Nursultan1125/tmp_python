from django.http.response import HttpResponse
import xlwt
from forms import *


class XLSBuilder:
    def __init__(self, id):
        self.id = id
        self.site_headers = [
            (u'ID', 2000),
            (u'Customer', 6000),
            (u'Customer Location', 6000),
            (u'Service Address', 6000),
            (u'City', 4000),
            (u'State', 2000),
            (u'Zip', 2000),
            (u'Site Use', 4000),
            (u'Site Type', 4000),
        ]
        self.survey_headers = [
            (u'ID', 2000),
            (u'Survey Date', 4000),
            (u'Surveyor', 4000),
            (u'Customer Rep', 4000),
            (u'Survey Type', 4000),
        ]
        self.finding_headers = [
            (u'NO.', 2000),
            (u'ID', 2000),
            (u'Building ID', 4000),
            (u'Room ID', 4000),
            (u'Location Description', 8000),
            (u'Latitude', 4000),
            (u'Longitude', 4000),
            (u'Description', 8000),
            (u'Degree of Hazard', 4000),
            (u'Regulation', 4000),
            (u'Subject to Backpressure', 4000),
            (u'CC Present', 4000),
            (u'Continuous Pressure', 4000),
            (u'Type BP Req\'d', 4000),
            (u'BP Present', 4000),
            (u'Size', 2000),
            (u'Type', 3000),
            (u'Installed Properly', 4000),
            (u'Type BP Present', 4000),
            (u'Manufacturer', 4000),
            (u'Model No.', 4000),
            (u'Serial No.', 4000),
            (u'Date Last Tested', 4000),
            (u'Installed Position', 4000),
            (u'Recommendations', 8000),
        ]
        self.col_pos, self.row_pos, self.findings_counter = 0, 2, 1

    def try_to_get(self, obj, field_name):
        try:
            return getattr(obj, field_name)
        except:
            return u'None'

    def site_data(self, obj):
        rows = [
            u'%s' % obj.id,
            u'%s' % obj.name,
            u'%s' % obj.location,
            u'%s' % obj.service_address,
            u'%s' % obj.city,
            u'%s' % obj.state,
            u'%s' % obj.zip,
            u'%s' % obj.site_use,
            u'%s' % self.try_to_get(obj, 'site_type'),
        ]
        return rows

    def survey_data(self, obj):
        try:
            surveyor_first_name = obj.surveyor.first_name
        except:
            surveyor_first_name = u''
        try:
            surveyor_last_name = obj.surveyor.last_name
        except:
            surveyor_last_name = u'None'
        rows = [
            u'%s' % obj.id,
            u'%s' % obj.date,
            u'%s %s' % (surveyor_first_name, surveyor_last_name),
            u'%s' % obj.customer_rep,
            u'%s' % self.try_to_get(obj, 'survey_type'),
        ]
        return rows

    def finding_data(self, obj):
        rows = [
            u'%s' % self.findings_counter,
            u'%s' % obj.id,
            u'%s' % obj.building_id,
            u'%s' % obj.room_id,
            u'%s' % obj.location_description,
            u'%s' % obj.latitude,
            u'%s' % obj.longitude,
            u'%s' % obj.description,
            u'%s' % self.try_to_get(obj, 'hazard_degree'),
            u'%s' % self.try_to_get(obj, 'regulation'),
            u'%s' % obj.subject_to_backpressure,
            u'%s' % obj.cc_present,
            u'%s' % obj.continuous_pressure,
            u'%s' % self.try_to_get(obj, 'bp_type_required'),
            u'%s' % obj.bp_present,
            u'%s' % self.try_to_get(obj, 'size'),
            u'%s' % self.try_to_get(obj, 'type'),
            u'%s' % obj.installed_properly,
            u'%s' % self.try_to_get(obj, 'bp_type'),
            u'%s' % self.try_to_get(obj, 'manufacturer'),
            u'%s' % obj.model_number,
            u'%s' % obj.serial_number,
            u'%s' % obj.last_tested_date,
            u'%s' % self.try_to_get(obj, 'installed_position'),
            u'%s' % obj.recommendation,
        ]
        return rows

    def build_headers(self, sheet, block_title, background_color, obj):
        self.row_pos = 0
        sheet.write_merge(self.row_pos, self.row_pos, self.col_pos, self.col_pos + len(obj) - 1, block_title,
                          xlwt.easyxf('font: bold 1; pattern: pattern solid, fore_colour %s;' % background_color))
        self.row_pos += 1
        for col_num in xrange(len(obj)):
            sheet.write(self.row_pos, self.col_pos + col_num, obj[col_num][0],
                        xlwt.easyxf('font: bold 1; pattern: pattern solid, fore_colour %s;' % background_color))
            sheet.col(self.col_pos + col_num).width = obj[col_num][1]
        self.col_pos += len(obj)

    def build_body(self, sheet, customer, survey, finding):
        for col_num in xrange(len(self.site_data(customer))):
            sheet.write(self.row_pos, col_num, self.site_data(customer)[col_num], xlwt.easyxf('alignment: wrap 1;'))
        self.col_pos = len(self.site_headers)
        if survey is not None:
            for col_num in xrange(len(self.survey_data(survey))):
                sheet.write(self.row_pos, self.col_pos + col_num, self.survey_data(survey)[col_num], xlwt.easyxf('alignment: wrap 1;'))
            self.col_pos += len(self.survey_headers)
            if finding is not None:
                for col_num in xrange(len(self.finding_data(finding))):
                    if self.finding_data(finding)[col_num]:
                        if self.finding_data(finding)[col_num] == u"True":
                            sheet.write(self.row_pos, self.col_pos + col_num, u"Yes", xlwt.easyxf('alignment: wrap 1;'))
                        elif self.finding_data(finding)[col_num] == u"False":
                            sheet.write(self.row_pos, self.col_pos + col_num, u"No", xlwt.easyxf('alignment: wrap 1;'))
                        else:
                            sheet.write(self.row_pos, self.col_pos + col_num, self.finding_data(finding)[col_num],
                                     xlwt.easyxf('alignment: wrap 1;'))
                    else:
                        sheet.write(self.row_pos, self.col_pos + col_num, u"None", xlwt.easyxf('alignment: wrap 1;'))
        self.row_pos += 1

    def survey_report(self):
        survey = Survey.objects.get(id=self.id)
        response = HttpResponse(mimetype='application/ms-excel')
        response['Content-Disposition'] = u'attachment; filename="%s_%s_report.xls"' % (survey.customer, survey.date)
        work_book = xlwt.Workbook(encoding='utf-8', style_compression=2)
        work_sheet = work_book.add_sheet(u'List1', cell_overwrite_ok=True)
        if survey.findings.all():
            for finding in survey.findings.all():
                self.build_body(work_sheet, survey.customer, survey, finding)
                self.findings_counter += 1
        else:
            self.build_body(work_sheet, survey.customer, survey, None)
        self.col_pos = 0
        self.build_headers(work_sheet, u'SITE INFORMATION', 'silver_ega', self.site_headers)
        self.build_headers(work_sheet, u'SURVEY INFORMATION', 'light_orange', self.survey_headers)
        self.build_headers(work_sheet, u'HAZARDS', 'white', self.finding_headers)
        work_book.save(response)
        return response

    def customer_report(self):
        customer = Customer.objects.get(id=self.id)
        response = HttpResponse(mimetype='application/ms-excel')
        response['Content-Disposition'] = u'attachment; filename="%s_%s_full_report.xls"' % (customer.name, customer.location)
        work_book = xlwt.Workbook(encoding='utf-8', style_compression=2)
        work_sheet = work_book.add_sheet(u'List1', cell_overwrite_ok=True)
        if customer.surveys.all():
            for survey in customer.surveys.all():
                self.findings_counter = 1
                if survey.findings.all():
                    for finding in survey.findings.all():
                        self.build_body(work_sheet, customer, survey, finding)
                        self.findings_counter += 1
                else:
                    self.build_body(work_sheet, customer, survey, None)
        else:
            self.build_body(work_sheet, customer, None, None)
        self.col_pos = 0
        self.build_headers(work_sheet, u'SITE INFORMATION', 'silver_ega', self.site_headers)
        self.build_headers(work_sheet, u'SURVEY INFORMATION', 'light_orange', self.survey_headers)
        self.build_headers(work_sheet, u'HAZARDS', 'white', self.finding_headers)
        work_book.save(response)
        return response
