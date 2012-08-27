<?php
/*
 * Google Spreadsheet Library
 * https://github.com/AlmirKadric/GoogleSpreadsheet-PHP
 *
 * Copyright 2012, Almir Kadric
 * https://almirkadric.com/
 *
 * Licensed under the MIT license:
 * http://www.opensource.org/licenses/MIT
 */
class GoogleSpreadsheet {
    private $token;
    private $spreadsheet;
    private $worksheet;
    private $cellLinks;
    private $listLinks;

    function __construct($username, $password) {
        if (!$this->authenticate($username, $password)) return false;
    }

    function authenticate($username, $password) {
        $this->token = null;
        $this->spreadsheet = null;
        $this->worksheet = null;
        $this->cellLinks = null;
        $this->listLinks = null;
        
        $url = "https://www.google.com/accounts/ClientLogin";

        $fields = array(
            "accountType" => "HOSTED_OR_GOOGLE",
            "Email" => $username,
            "Passwd" => $password,
            "service" => "wise",
            "source" => "pfbc"
        );

        if ($response = $this->post($url, false, $fields)) {
            if (stripos($response, "auth=") !== false) {
                $matches = array();
                preg_match("/auth=([a-z0-9_\-]+)/i", $response, $matches);
                $this->token = $matches[1];
            }
            
            return true;
        }
        
        echo "\nError Authenticating User!\n";
        return false;
    }
    
    function listSpreadsheets()
    {
        
    }
    
    function listWorksheets()
    {
        if (empty($this->token) || empty($this->spreadsheet)) return false;
        
        $url = $this->spreadsheet['links']['content'];
        
        $headers = array(
            "Content-Type: application/atom+xml",
            "Authorization: GoogleLogin auth=" . $this->token,
            "GData-Version: 3.0"
        );
        
        if ($response = $this->post($url, $headers))
        {
            $responseXml = simplexml_load_string($response);
            $responseNss = $responseXml->getNamespaces(true);
            
            $worksheets = array();
            for ($i = 0; $i < $responseXml->entry->count(); $i++) $worksheets[] = (string)$responseXml->entry[$i]->title;
            
            return $worksheets;
        } else {
            echo "\nError retrieving worksheets list!\n";
            return false;
        }
    }
    
    function createSpreadsheet($title)
    {
        echo "\nError creating spreadsheet!\n";
        echo "\nNOT IMPLEMENTED YET!\n";
        return false;
    }
    
    function createWorksheet($title, $columnNames)
    {
        if (empty($this->token) || empty($this->spreadsheet)) return false;
        
        $url = $this->spreadsheet['links']['content'];
        
        $headers = array(
            "Content-Type: application/atom+xml",
            "Authorization: GoogleLogin auth=" . $this->token,
            "GData-Version: 3.0"
        );
        
        $xml  = "<entry xmlns=\"http://www.w3.org/2005/Atom\" xmlns:gs=\"http://schemas.google.com/spreadsheets/2006\">" . "\n";
        $xml .= "    <title>" . $title . "</title>" . "\n";
        $xml .= "    <gs:rowCount>1</gs:rowCount>" . "\n";
        $xml .= "    <gs:colCount>" . count($columnNames) . "</gs:colCount>" . "\n";
        $xml .= "</entry>" . "\n";
        
        // If couldn't create return error
        if (!$response = $this->post($url, $headers, $xml))
        {
            echo "\nError creating new worksheet!\n";
            return false;
        }
        
        // Otherwise set the worksheet and insert column names
        $this->useWorksheet($title);
        foreach ($columnNames as $i => $columnName) $this->updateCell(1, ($i + 1), $columnName);
        
        return true;
    }
    
    function useSpreadsheet($title) {
        $this->spreadsheet = array();
        $this->worksheet = null;
        $this->cellLinks = null;
        $this->listLinks = null;
        
        if (empty($this->token)) return false;

        $url = "https://spreadsheets.google.com/feeds/spreadsheets/private/full?title=" . urlencode($title);

        $headers = array(
                "Authorization: GoogleLogin auth=" . $this->token,
                "GData-Version: 3.0"
        );

        if ($response = $this->post($url, $headers)) {
            $spreadsheetXml = simplexml_load_string($response);
            $spreadsheetNss = $spreadsheetXml->getNamespaces(true);

            if ($spreadsheetXml->entry) {
                $this->spreadsheet['id'] = basename(trim($spreadsheetXml->entry[0]->id));
                $this->spreadsheet['title'] = (string)$spreadsheetXml->entry[0]->title;
                $this->spreadsheet['updated'] = (string)$spreadsheetXml->entry[0]->updated;

                $this->spreadsheet['author'] = array(
                    'name' => (string)$spreadsheetXml->entry[0]->author->name,
                    'email' => (string)$spreadsheetXml->entry[0]->author->email
                );

                $this->spreadsheet['links']['content'] = (string)$spreadsheetXml->entry[0]->content->attributes()->src;
                for ($i = 0; $i < $spreadsheetXml->entry[0]->link->count(); $i++) {
                    $rel = (string)$spreadsheetXml->entry[0]->link[$i]->attributes()->rel;
                    $href = (string)$spreadsheetXml->entry[0]->link[$i]->attributes()->href;
                    $this->spreadsheet['links'][$rel] = $href;
                }
                
                return true;
            } else {
                return $this->createSpreadsheet($title, $columnNames);
            }
        }
        
        echo "\nError setting spreadsheet!\n";
        return false;
    }
    
    function useWorksheet($title, $columnNames = false) {
        $this->worksheet = array();
        $this->cellLinks = null;
        $this->listLinks = null;
        
        if (empty($this->token) || empty($this->spreadsheet)) return false;

        $url = $this->spreadsheet['links']['content'] . "?title=" . urlencode($title);

        $headers = array(
                "Authorization: GoogleLogin auth=" . $this->token,
                "GData-Version: 3.0"
        );

        if ($response = $this->post($url, $headers)) {
            $worksheetXml = simplexml_load_string($response);
            $worksheetNss = $worksheetXml->getNamespaces(true);

            if ($worksheetXml->entry) {
                $this->worksheet['id'] = basename(trim($worksheetXml->entry[0]->id));
                $this->worksheet['title'] = (string)$worksheetXml->entry[0]->title;
                $this->worksheet['updated'] = (string)$worksheetXml->entry[0]->updated;

                $this->worksheet['links']['content'] = (string)$worksheetXml->entry[0]->content->attributes()->src;
                for ($i = 0; $i < $worksheetXml->entry[0]->link->count(); $i++) {
                    $rel = (string)$worksheetXml->entry[0]->link[$i]->attributes()->rel;
                    $href = (string)$worksheetXml->entry[0]->link[$i]->attributes()->href;
                    $this->worksheet['links'][$rel] = $href;
                }

                $gsChildren = $worksheetXml->entry[0]->children($worksheetNss['gs']);
                $this->worksheet['rows'] = (string)$gsChildren->rowCount - 1;
                $this->worksheet['columns'] = (string)$gsChildren->colCount;
                
                return true;
            } else {
                if ($columnNames) {
                    return $this->createWorksheet($title, $columnNames);
                } else {
                    echo "\nCould not auto-create worksheet!\n";
                    echo "\nMust specify column names!\n";
                    return false;
                }
            }
        }
        
        echo "\nError setting worksheet!\n";
        return false;
    }
    
    function getCellUrls() {
        $this->cellLinks = array();
        
        if (empty($this->token) || empty($this->worksheet)) return false;
        
        $url = $this->worksheet['links']['http://schemas.google.com/spreadsheets/2006#cellsfeed'] . "?q=''";

        $headers = array(
                "Authorization: GoogleLogin auth=" . $this->token,
                "GData-Version: 3.0"
        );

        if ($response = $this->post($url, $headers)) {
            $cellfeedXml = simplexml_load_string($response);
            $cellfeedNss = $cellfeedXml->getNamespaces(true);

            for ($i = 0; $i < $cellfeedXml->link->count(); $i++) {
                $rel = (string)$cellfeedXml->link[$i]->attributes()->rel;
                $href = (string)$cellfeedXml->link[$i]->attributes()->href;
                $this->cellLinks[$rel] = $href;
            }
        }
    }
    
    function getListUrls() {
        $this->listLinks = array();
        
        if (empty($this->token) || empty($this->worksheet)) return false;
        
        $url = $this->worksheet['links']['content'] . "?q=''";

        $headers = array(
                "Authorization: GoogleLogin auth=" . $this->token,
                "GData-Version: 3.0"
        );

        if ($response = $this->post($url, $headers)) {
            $listfeedXml = simplexml_load_string($response);
            $listfeedNss = $listfeedXml->getNamespaces(true);

            for ($i = 0; $i < $listfeedXml->link->count(); $i++) {
                $rel = (string)$listfeedXml->link[$i]->attributes()->rel;
                $href = (string)$listfeedXml->link[$i]->attributes()->href;
                $this->listLinks[$rel] = $href;
            }
        }
    }
    
    function getCellFeedUrl() 
    {
        if (empty($this->cellLinks)) $this->getCellUrls();
        
        return $this->cellLinks['http://schemas.google.com/g/2005#feed'];
    }
    
    function getCellBatchUrl() 
    {
        if (empty($this->cellLinks)) $this->getCellUrls();
        
        return $this->cellLinks['http://schemas.google.com/g/2005#batch'];
    }
    
    function getCellPostUrl() 
    {
        if (empty($this->cellLinks)) $this->getCellUrls();
        
        return $this->cellLinks['http://schemas.google.com/g/2005#post'];
    }
    
    function getListFeedUrl() 
    {
        if (empty($this->listLinks)) $this->getListUrls();
        
        return $this->listLinks['http://schemas.google.com/g/2005#feed'];
    }
    
    function getListPostUrl() 
    {
        if (empty($this->listLinks)) $this->getListUrls();
        
        return $this->listLinks['http://schemas.google.com/g/2005#post'];
    }
    
    function getColumnNames() {
        if (empty($this->token) || empty($this->worksheet)) return false;
        
        $url = $this->getCellFeedUrl() . "?min-row=1&max-row=1";

        $headers = array(
            "Authorization: GoogleLogin auth=" . $this->token,
            "GData-Version: 3.0"
        );

        if ($response = $this->post($url, $headers)) {
            $columnXml = simplexml_load_string($response);
            $columnNss = $columnXml->getNamespaces(true);

            $columns = array();
            for ($i = 0; $i < $columnXml->entry->count(); $i++) {
                $gsChildren = $columnXml->entry[$i]->children($columnNss['gs']);

                $columnName = (string)$gsChildren[0]->attributes()->inputValue;
                $columnNumber = (string)$gsChildren[0]->attributes()->col;

                $columns[$columnName] = $columnNumber;
                $columns[$columnNumber] = $columnName;
            }

            return $columns;
        }
        
        echo "\nError grabbing column names!\n";
        return false;
    }
    
    function getRowCount()
    {
        if (empty($this->token) || empty($this->worksheet)) return false;
        
        $url = $this->worksheet['links']['self'];
        
        $headers = array(
            "Content-Type: application/atom+xml",
            "Authorization: GoogleLogin auth=" . $this->token,
            "GData-Version: 3.0"
        );
        
        if ($response = $this->post($url, $headers)) {
            $returnXml = simplexml_load_string($response);
            $returnNss = $returnXml->getNamespaces(true);
            
            $gsChildren = $returnXml->children($returnNss['gs']);
            return (string)$gsChildren->rowCount;
        } else {
            echo "\nError retrieving row count!\n";
            return false;
        }
    }
    
    function getCellRows($startRow = null, $endRow = null, $startCol = null, $endCol = null)
    {
        if (empty($this->token) || empty($this->worksheet)) return false;
        
        $url = $this->getCellFeedUrl() . "?";
        if ($startRow !== null) $url .= "min-row=$startRow&";
        if ($endRow !== null) $url .= "max-row=$endRow&";
        if ($startCol !== null) $url .= "min-col=$startCol&";
        if ($endCol !== null) $url .= "max-col=$endCol&";
        
        $headers = array(
            "Content-Type: application/atom+xml",
            "Authorization: GoogleLogin auth=" . $this->token,
            "GData-Version: 3.0"
        );
        
        if ($response = $this->post($url, $headers)) {
            $rowsXml = simplexml_load_string($response);
            $rowsNss = $rowsXml->getNamespaces(true);
            
            $rows = array();
            for ($i = 0; $i < $rowsXml->entry->count(); $i++) {
                $gsChildren = $rowsXml->entry[$i]->children($rowsNss['gs']);
                
                $row = (string)$gsChildren[0]->attributes()->row;
                $column = (string)$gsChildren[0]->attributes()->col;
                $value = (string)$rowsXml->entry[$i]->content;
                
                $rows[$row][$column] = $value;
            }
            
            return $rows;
        } else {
            echo "\nError retrieving rows!\n";
            return false;
        }
    }
    
    function insertRow($row)
    {
        if (empty($this->token) || empty($this->worksheet)) return false;
        
        $url = $this->getListPostUrl();
        
        $headers = array(
            "Content-Type: application/atom+xml",
            "Authorization: GoogleLogin auth=" . $this->token,
            "GData-Version: 3.0"
        );
        
        $xml = "<entry xmlns=\"http://www.w3.org/2005/Atom\" xmlns:gsx=\"http://schemas.google.com/spreadsheets/2006/extended\">" . "\n";
        
        foreach ($row as $key => $value) {
            $key = strtolower($key);
            $xml .= "<gsx:$key>$value</gsx:$key>" . "\n";
        }
        
        $xml .= "</entry>" . "\n";
        
        if ($response = $this->post($url, $headers, $xml)) {
            $returnXml = simplexml_load_string($response);
            $returnNss = $returnXml->getNamespaces(true);
            
            // TODO NEED TO HANDLE BUGGY GOOGLE RESPONSES
            
            // Additional section to handle formulas within the insert
            //$this->updateFormulas();
        } else {
            echo "\nError inserting new row into worksheet!\n";
            return false;
        }
        
        return true;
    }
    
    function updateCell($row, $column, $value)
    {
        if (empty($this->token) || empty($this->worksheet)) return false;
        
        $url = $this->getCellPostUrl();
        
        $headers = array(
            "Content-Type: application/atom+xml",
            "Authorization: GoogleLogin auth=" . $this->token,
            "GData-Version: 3.0"
        );
        
        $xml = "<entry xmlns=\"http://www.w3.org/2005/Atom\" xmlns:gs=\"http://schemas.google.com/spreadsheets/2006\">" . "\n";
            $xml .= "<id>" . $this->getCellFeedUrl() . "/" . $this->generateAddress($row, $column, false, false) . "</id>" . "\n";
            $xml .= "<link rel=\"edit\" type=\"application/atom+xml\" href=\"" . $this->getCellFeedUrl() . "/" . $this->generateAddress($row, $column, false, false) . "\"/>" . "\n";
            $xml .= "<gs:cell row=\"" . $row . "\" col=\"" . $column . "\" inputValue=\"" . htmlentities($value) . "\"/>" . "\n";
        $xml .= "</entry>" . "\n";
        
        if ($response = $this->post($url, $headers, $xml)) {
            $returnXml = simplexml_load_string($response);
            $returnNss = $returnXml->getNamespaces(true);

            // TODO NEED TO HANDLE BUGGY GOOGLE RESPONSES
        } else {
            echo "\nError updating cell data! (".$this->generateAddress($row, $column).")\n";
            return false;
        }
        
        return true;
    }
    
    function updateCells($data)
    {
        if (empty($this->token) || empty($this->worksheet)) return false;
        
        $batchUrl = $this->getCellBatchUrl();

        $headers = array(
            "Content-Type: application/atom+xml",
            "Authorization: GoogleLogin auth=" . $this->token,
            "GData-Version: 3.0",
            "If-Match: *"
        );

        $xml = "<feed xmlns=\"http://www.w3.org/2005/Atom\"" . "\n" .
                    "xmlns:batch=\"http://schemas.google.com/gdata/batch\"" . "\n" .
                    "xmlns:gs=\"http://schemas.google.com/spreadsheets/2006\">" . "\n";

        $xml .= "<id>" . $batchUrl . "</id>" . "\n";
        
        foreach ($data as $row => $entry) {
            foreach ($entry as $column => $value) {
                $xml .= "<entry>" . "\n";

                    $xml .= "<batch:id>" . $this->generateAddress($row, $column) . "</batch:id>" . "\n";
                    $xml .= "<batch:operation type=\"update\"/>" . "\n";
                    $xml .= "<title type=\"text\">" . $this->generateAddress($row, $column) . "</title>" . "\n";
                    $xml .= "<id>" . $this->getCellFeedUrl() . "/" . $this->generateAddress($row, $column, false, false) . "</id>" . "\n";
                    $xml .= "<link rel=\"edit\" type=\"application/atom+xml\" href=\"" . $this->getCellFeedUrl() . "/" . $this->generateAddress($row, $column, false, false) . "/0" . "\"/>" . "\n";
                    $xml .= "<gs:cell row=\"" . $row . "\" col=\"" . $column . "\" inputValue=\"" . htmlentities($value) . "\"/>" . "\n";

                $xml .= "</entry>" . "\n";
            }
        }

        $xml .= "</feed>" . "\n";
        
        if ($response = $this->post($batchUrl, $headers, $xml)) {
            $returnXml = simplexml_load_string($response);
            $returnNss = $returnXml->getNamespaces(true);

            // TODO NEED TO HANDLE BUGGY GOOGLE RESPONSES
        } else {
            echo "\nError saving data to worksheet! (".$this->generateAddress($row, $column).")\n";
            return false;
        }

        return true;
    }
    
    function deleteRow($row)
    {
        if (empty($this->token) || empty($this->worksheet)) return false;
        
        $columns = $this->getColumnNames();
        
        $query = array();
        foreach($row as $col => $value) $query[] = strtolower($columns[$col]) . " = \"" . $value . "\"";
        $query = urlencode(implode(" AND ", $query));
        
        $url = $this->getListFeedUrl() . "?sq=$query";
        
        $headers = array(
            "Content-Type: application/atom+xml",
            "Authorization: GoogleLogin auth=" . $this->token,
            "GData-Version: 3.0",
            "If-Match: *"
        );
        
        if ($response = $this->post($url, $headers)) {
            $rowXml = simplexml_load_string($response);
            
            $deleteUrl = (string)$rowXml->entry[$rowXml->entry->count() - 1]->link->attributes()->href;
            
            if (($response = $this->post($deleteUrl, $headers, null, "DELETE", $status)) === false) {
                echo "\nError deleting row!\n";
                return false;
            }
            
            return true;
        } else {
            echo "\nError retrieving row during delete operation!\n";
            return false;
        }
    }
    
    function post($url, $headers = false, $post = false, $request = null, &$status = null, &$error = null, $retryCount = 0)
    {
        $curl = curl_init();
        
        curl_setopt($curl, CURLOPT_URL, $url);
        curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false);
        curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
        
        if ($headers) curl_setopt($curl, CURLOPT_HTTPHEADER, $headers);
        
        if ($post) {
            curl_setopt($curl, CURLOPT_POST, true);
            curl_setopt($curl, CURLOPT_POSTFIELDS, $post);
        }
        
        if ($request) {
            curl_setopt($curl, CURLOPT_CUSTOMREQUEST, $request);
        }
        
        $response = curl_exec($curl);
        $status = curl_getinfo($curl, CURLINFO_HTTP_CODE);
        
        curl_close($curl);
        
        // If Successfull return results
        if(in_array($status, array('200', '201'))) return $response;
        
        // If server error, keepy trying
        if (in_array($status, array('400'))) {
            echo html_entity_decode($response) . "\n";
            echo "Retrying in 5 seconds...\n\n";
            
            sleep(30);
            if ($retryCount < 60) {
                $retryCount++;
                return $this->post($url, $headers, $post, $request, $status, $error, $retryCount);
            }
        }
        
        // Failed, return an error
        $error = $response;
        
        echo $error . "\n";
        echo "\nStatus: $status\n";
        
        return false;
    }
    
    function xmlPretty($xml)
    {
        $xml = str_replace(array(">", "</"), array(">\n", "\n</"), $xml);
        
        $xml = explode("\n", $xml);
        $tabIndex = 0;
        $tabSize = 4;
        foreach ($xml as $i => &$string) {
            // Remove empty lines
            if ($string == "") {
                unset($xml[$i]);
                continue;
            }
            
            // Shorten indent size if closing element
            if (preg_match("/^<\//", trim($string))) $tabIndex--;
            
            // Indent string accordingly
            for ($j = 0; $j < ($tabIndex * $tabSize); $j++) $string = " " . $string;
            
            // Increase indent size if opening element
            if (preg_match("/^</", trim($string)) && !preg_match("/^<\?|^<\/|\/>$/", trim($string))) $tabIndex++;
            
            // Make single value elements 1 liners
            if (preg_match("/^<\//", trim($string))) {
                if (isset($xml[$i - 1]) && isset($xml[$i - 1])) {
                    if (preg_match("/^</", trim($xml[$i - 2])) && !preg_match("/^<\?|^<\/|\/>$/", trim($xml[$i - 2]))) {
                        if (!preg_match("/^</", trim($xml[$i - 1]))) {
                            $xml[$i - 2] .= trim($xml[$i - 1]) . trim($string);
                            unset($xml[$i - 1]);
                            unset($xml[$i]);
                        }
                    }
                }
            }
            
            // Decode URL html entities
            $string = html_entity_decode($string);
        }
        
        return implode("\n", $xml);
    }
    
    function generateColAplha($number)
    {
        if ($number > 26) return "Z" . $this->generateColAplha($number - 26);
        else return chr(64 + $number);
    }
    
    function generateAddress($row, $column, $absolute = false, $reference = true)
    {
        if ($reference) return (($absolute)?"$":"") . $this->generateColAplha($column) . (($absolute)?"$":"") . $row;
        else return "R" . $row . "C" . $column;
    }
}
?>