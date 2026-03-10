"""
Format definitions for the Summary sheet.

This module centralizes all format objects used in the Summary sheet builder,
keeping format definitions separate from sheet layout and data logic.
"""


def create_summary_formats(workbook):
    """
    Create all format objects for the Summary sheet.

    Args:
        workbook: xlsxwriter.Workbook instance

    Returns:
        dict: Dictionary of format objects keyed by format name
    """
    formats = {}

    # Title and section formats
    formats['fmt_title'] = workbook.add_format({
        'font_size': 18, 'bold': True,
        'align': 'center_across'
    })

    formats['fmt_subtitle'] = workbook.add_format({
        'font_size': 14, 'bold': True,
        'align': 'center_across'
    })

    formats['fmt_section'] = workbook.add_format({
        'font_size': 10, 'bold': True,
        'align': 'center_across'
    })

    formats['fmt_section_left'] = workbook.add_format({
        'font_size': 10, 'bold': True
    })

    # Header formats
    formats['fmt_hdr_white_center_top'] = workbook.add_format({
        'font_size': 10, 'bold': True, 'italic': True,
        'align': 'center_across', 'bottom': 1, 'top': 1
    })

    formats['fmt_hdr_white_border'] = workbook.add_format({
        'font_size': 10, 'bold': True,
        'align': 'vcenter', 'bottom': 1
    })

    # Label formats
    formats['fmt_label'] = workbook.add_format({
        'font_size': 10
    })

    formats['fmt_label_border'] = workbook.add_format({
        'font_size': 10, 'bottom': 1
    })

    formats['fmt_sub_label'] = workbook.add_format({
        'font_size': 9
    })

    # Data value formats
    formats['fmt_impl_white'] = workbook.add_format({
        'font_size': 10, 'bold': True
    })

    # Border formats for the outside box
    formats['fmt_border'] = workbook.add_format({'bottom': 2})
    formats['fmt_border_thin'] = workbook.add_format({'bottom': 1})
    formats['fmt_border_top'] = workbook.add_format({'top': 1})
    formats['fmt_border_bottom'] = workbook.add_format({'bottom': 1})
    formats['fmt_border_left'] = workbook.add_format({'left': 1})
    formats['fmt_border_right'] = workbook.add_format({'right': 1})
    formats['fmt_border_top_left'] = workbook.add_format({'top': 1, 'left': 1})
    formats['fmt_border_top_right'] = workbook.add_format({'top': 1, 'right': 1})
    formats['fmt_border_bottom_left'] = workbook.add_format({'bottom': 1, 'left': 1})
    formats['fmt_border_bottom_right'] = workbook.add_format({'bottom': 1, 'right': 1})

    # Separator formats
    formats['fmt_dashed_border'] = workbook.add_format({'bottom': 4})

    return formats
