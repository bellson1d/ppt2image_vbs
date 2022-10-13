Function FormatNumberLeadingZeroes(expr, digits)
    remain = expr

    Do 
      result = CStr(remain mod 10) & result
      remain = remain \ 10
    Loop Until (remain = 0) and (Len(result) >= digits)

  FormatNumberLeadingZeroes = result
End Function


function ExportSlides(ppt_file, out_format, megapixels)

  Set pptApp = CreateObject("PowerPoint.Application")
  Set ppt = pptApp.Presentations.Open(ppt_file, True, , False)

  With ppt.PageSetup
    sh = .SlideHeight
    sw = .SlideWidth
  End With

  sA = sh * sw
  factor = Sqr(1000000.0 * megapixels / sA)
  imageheight = Round(factor * sh, 0)
  imagewidth = Round(factor * sw, 0)
  num_exported = 0

  For Each slide in ppt.Slides

    slide.Export ppt_file & "-" _
        & FormatNumberLeadingZeroes(slide.SlideIndex, 4) _
        & "." & LCase(out_format), _
        out_format, _
        imagewidth, imageheight

    num_exported = num_exported + 1

  Next

  ExportSlides = num_exported
End Function

For Each arg in Wscript.Arguments 
  ' adjust megapixels here in the last argument:
  num_exported = ExportSlides(arg, "JPG", 4)
  WScript.Echo "Done with ", arg, ". ", num_exported, " slides exported."
Next