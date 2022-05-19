from docxtpl import DocxTemplate
import docx.opc.exceptions
from alert_windows import no_mail, no_act, no_plan, successful
from vsdx import VisioFile
import os.path

doc_mail = DocxTemplate('письмо_шаблон.docx')
doc_act = DocxTemplate('акт_шаблон.docx')
doc_plan = DocxTemplate('схема_шаблон.docx')


# создание письма
def mail_save(context):
    try:
        doc_mail.render(context)
        doc_mail.save(f'письмо {context["file_name"]}.docx')
        if os.path.isfile(f'письмо {context["file_name"]}.docx'):
            successful(f'письмо {context["file_name"]}.docx')
    except docx.opc.exceptions.PackageNotFoundError:
        no_mail()

# создание акта
def act_save(context):
    try:
        doc_act.render(context)
        doc_act.save(f'акт {context["file_name"]}.docx')
        if os.path.isfile(f'акт {context["file_name"]}.docx'):
            successful(f'акт {context["file_name"]}.docx')
    except docx.opc.exceptions.PackageNotFoundError:
        no_act()

# создание схемы
def plan_save(context):
    try:
        with VisioFile('схема_шаблон.vsdx') as vis:
            page = vis.pages[0]
            page.apply_text_context(context)
            vis.save_vsdx(f"схема {context['file_name']}.vsdx")
            if os.path.isfile(f'схема {context["file_name"]}.vsdx'):
                successful(f'схема {context["file_name"]}.vsdx')
    except FileNotFoundError:
        no_plan()
