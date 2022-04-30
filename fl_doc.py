from docxtpl import DocxTemplate
from fl_def import get_file_name

doc_mail = DocxTemplate('письмо_шаблон.docx')
doc_act = DocxTemplate('акт_шаблон.docx')
doc_plan = DocxTemplate('схема_шаблон.docx')

# создание письма
def mail_save(context):
    doc_mail.render(context)
    file_name = get_file_name(context['organization'])
    doc_mail.save(f'письмо {file_name}.docx')

# создание акта
def act_save(context):
    doc_act.render(context)
    file_name = get_file_name(context['organization'])
    doc_act.save(f'акт {file_name}.docx')

# создание схемы
def plan_save(context):
    doc_plan.render(context)
    file_name = get_file_name(context['organization'])
    doc_plan.save(f'схема {file_name}.docx')