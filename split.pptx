import os
import win32com.client
import time

def split_pptx_win32com(input_path, slides_per_part, output_dir="."):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True

    presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
    total_slides = presentation.Slides.Count

    for i in range(0, total_slides, slides_per_part):
        part_ppt = powerpoint.Presentations.Add()
        for j in range(i + 1, min(i + slides_per_part + 1, total_slides + 1)):
            presentation.Slides(j).Copy()
            time.sleep(0.2)
            part_ppt.Slides.Paste()

        part_number = i // slides_per_part + 1
        part_filename = os.path.join(output_dir, f"part{part_number}.pptx")
        part_ppt.SaveAs(part_filename)
        print(f"Zapisano: {part_filename}")
        part_ppt.Close()

    presentation.Close()
    powerpoint.Quit()


split_pptx_win32com(r"A:\prezka\7_ava\all\6.pptx", slides_per_part=169, output_dir=r"A:\prezka\7_ava\all\\")
