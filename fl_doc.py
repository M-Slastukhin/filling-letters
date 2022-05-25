from docxtpl import DocxTemplate
import docx.opc.exceptions
from alert_windows import no_docx, no_plan, successful, bad_xlsx
from vsdx import VisioFile
import os.path


# создание docx файла
def docx_save(context, sample_docx, doc_form):
    try:
        docx_file = DocxTemplate(sample_docx)
        docx_file.render(context)
        docx_file.save(f'{doc_form} {context["file_name"]}.docx')
        if os.path.isfile(f'{doc_form} {context["file_name"]}.docx'):
            successful(f'{doc_form} {context["file_name"]}.docx')
    except docx.opc.exceptions.PackageNotFoundError:
        no_docx()
    except ValueError:
        no_docx()
    except KeyError:
        bad_xlsx()

# создание vsdx файла
def vsdx_save(context, sample_vsdx):
    try:
        with VisioFile(sample_vsdx) as vis:
            page = vis.pages[0]
            page.apply_text_context(context)
            vis.save_vsdx(f"схема {context['file_name']}.vsdx")
            if os.path.isfile(f'схема {context["file_name"]}.vsdx'):
                successful(f'схема {context["file_name"]}.vsdx')
    except TypeError:
        no_plan()
    except KeyError:
        bad_xlsx()