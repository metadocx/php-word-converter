<?php

use Illuminate\Support\Facades\Route;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Log;
use Metadocx\Reporting\Converters\Word\WordConverter;

Route::post("/Metadocx/Convert/Word", function(Request $request) {

    $oWordConverter = new WordConverter();    
    $oWordConverter->loadOptions($request->input("ExportOptions"));    
    $sFileName = $oWordConverter->convert($request->input("HTML"));
    if ($sFileName !== false) {       
        $headers = ["Content-Type"=> "application/octet-stream"];
        return response()
                ->download($sFileName, "Report.docx", $headers)
                ->deleteFileAfterSend(true);
    } 
});