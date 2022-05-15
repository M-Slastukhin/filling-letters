from docxtpl import DocxTemplate
import docx.opc.exceptions
from fl_def import get_file_name
from alert_windows import no_mail, no_act, no_plan
from vsdx import VisioFile

doc_mail = DocxTemplate('письмо_шаблон.docx')
doc_act = DocxTemplate('акт_шаблон.docx')
doc_plan = DocxTemplate('схема_шаблон.docx')


# создание письма
def mail_save(context):
    try:
        doc_mail.render(context)
        file_name = get_file_name(context['organization'])
        doc_mail.save(f'письмо {file_name}.docx')
    except docx.opc.exceptions.PackageNotFoundError:
        no_mail()

# создание акта
def act_save(context):
    try:
        doc_act.render(context)
        file_name = get_file_name(context['organization'])
        doc_act.save(f'акт {file_name}.docx')
    except docx.opc.exceptions.PackageNotFoundError:
        no_act()

# создание схемы
def plan_save(context):
    try:
        file_name = get_file_name(context['organization'])
        with VisioFile('схема_шаблон.vsdx') as vis:
            page = vis.pages[0]
            page.apply_text_context(context)
            vis.save_vsdx(f'схема {file_name}.vsdx')
    except FileNotFoundError:
        no_plan()
     #   save_vsdx(new_filename=f'схема {file_name}')
    #doc_plan.render(context)
    #file_name = get_file_name(context['organization'])
    #doc_plan.save(f'схема {file_name}.docx')