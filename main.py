"""Cloud Functions entry point for Excel-PowerPoint Merger."""

import functions_framework


@functions_framework.http
def excel_pptx_merger(request):
    """Cloud Function entry point - use the original working implementation."""
    # Import and use the original, fully-functional Cloud Functions handler
    from src.main import excel_pptx_merger as original_handler
    return original_handler(request)
