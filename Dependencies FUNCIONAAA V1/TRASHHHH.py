
def update_charts_in_pptxDEPRECATED2(ppt_export_location: str, preguntas_parseadas) -> str:
    """
    Pipeline:
      1) Unzip PPTX into a working folder.
      2) List all chart XML files under ppt/charts.
      3) Pair one-to-one with preguntas_parseadas.
      4) Apply XML updates per pair.
      5) Repack into '*_updated.pptx'.
    """
    tmp_dir = unzip_pptx_clean(ppt_export_location)
    chart_files = list_chart_files(tmp_dir)
    pairs = match_charts_to_questions(chart_files, preguntas_parseadas)

    for chart_path, q in pairs:
        print(f"Updating {os.path.basename(chart_path)} for question: {q.nombre}")
        update_chart_for_question(chart_path, q)

    output_path = ppt_export_location.replace(".pptx", "_updated.pptx")
    rezip_pptx(tmp_dir, output_path)

    # Optionally refresh all graphs after rezip:
    # refresh_all_graphs(output_path)

    print(f"New file created: {output_path}")
    return output_path

def _all_series_must_live_in_barchart(root) -> None:
    """
    If there are stray <c:ser> directly under <c:plotArea>, move them into
    <c:barChart> (if present). Does not create/delete series.
    """
    bar = root.find(".//c:barChart", NS)
    if bar is None:
        return
    plot_area = root.find(".//c:plotArea", NS)
    for ser in list(plot_area.findall("./c:ser", NS)):
        plot_area.remove(ser)
        bar.append(ser)