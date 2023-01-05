<?php 
namespace Metadocx\Reporting\Converters\Word;

use Illuminate\Support\Facades\Storage;
use Illuminate\Support\Facades\Log;

class WordConverter {
    
    protected $_sOutputFileName = null;
    protected $_aOptions = [];

    public function convert($sContent) {
        

        $oWordDocument = new \PhpOffice\PhpWord\PhpWord();
        $oDocumentSection = $oWordDocument->addSection();
     
        $doc = new \DOMDocument();
        libxml_use_internal_errors(true);
        $doc->loadHTML(base64_decode($sContent), LIBXML_NOERROR);
        $doc->normalizeDocument();

        \PhpOffice\PhpWord\Shared\Html::addHtml($oDocumentSection, $doc->saveHTML(), true);

        $this->_sOutputFileName = storage_path("app/" . uniqid("Word") . ".docx");
        $oWordDocument->save($this->_sOutputFileName, 'Word2007');

        if (file_exists($this->_sOutputFileName)) {

            return $this->_sOutputFileName;
        } else {
            return false;
        }

    }

    public function loadOptions($options) {

        $this->_aOptions = [];

    }

    public function __get(string $name) {
        $name = str_replace("_", "-", $name);
        return $this->_aOptions[$name];
    }

    public function __set(string $name, mixed $value) {
        $name = str_replace("_", "-", $name);
        $this->_aOptions[$name] = $value;
        return $this;
    }

    private function toBool($value) {
        if (is_bool($value)) {
            return (bool) $value;
        }

        $value = strtolower(trim($value));

        $aTrueValues = ["1","y","o","yes","true","oui","vrai","on","checked",true,1];
        
        if (in_array($value, $aTrueValues, true)) {
            return true;
        } else {
            return false;
        }
    }

}