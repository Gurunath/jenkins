<?php

/*
  Cron
  Created Date : 21 Nov 2018
  Created By : Gurunath Auti
  Modified By : Gurunath Auti 16 Nov 2018
  Purpose : Send a weekly call report acrding to agency client weekly format and Call Details .
  It will be run on every sunday
 */


require_once '/usr1/htdocs/erp/bs/assets/libmail/libmail.php';
require_once '/usr1/htdocs/erp/bs/assets/library/PHPExcel/Classes/PHPExcel.php';

$firstDay = date('Y-m-01');
$lastDay = date('Y-m-t');

#$firstDay = date('2018-11-01');
#$lastDay = date('2018-11-30');

$con = mysqli_connect('sitedbserver', 'erp', 'erp123', 'erpapp') or die("error in connection");
#mysql_select_db('erpapp');

echo $query = "SELECT  userid, l.branch, executiveData.username,  agtype,
dateTime, WEEK(dateTime,6) - WEEK(DATE_SUB(dateTime, INTERVAL DAYOFMONTH(dateTime)-1 DAY),6)+1 as week,
 count(dateTime) as dayCall
            FROM erpapp.executiveData
            LEFT JOIN login l ON l.id = executiveData.userid
            WHERE dateTime between '" . $firstDay . "' AND '" . $lastDay . "'   
                AND l.dept = 'Web-Business'
            group by WEEK(dateTime), agtype, userid
            order by l.branch,dateTime desc";

$result = mysqli_query($con, $query);

while ($res = mysqli_fetch_array($result)) {
    $arrExeCalls[] = $res;
}

$arrUserCall = array();

$arrWeek = array('1AGENCY', '2AGENCY', '3AGENCY', '4AGENCY', '1DIRECT', '2DIRECT', '3DIRECT', '4DIRECT');

foreach ($arrExeCalls as $key => $exeCall) {
    foreach ($arrWeek as $week) {

        $strWeek = $exeCall['week'] . $exeCall['agtype'];

        $arrUserCall[$exeCall['userid']][$exeCall['week'] . $exeCall['agtype']]['username'] = $exeCall['username'];
        $arrUserCall[$exeCall['userid']][$exeCall['week'] . $exeCall['agtype']]['branch'] = $exeCall['branch'];
        $arrUserCall[$exeCall['userid']][$exeCall['week'] . $exeCall['agtype']]['userid'] = $exeCall['userid'];
        $arrUserCall[$exeCall['userid']][$exeCall['week'] . $exeCall['agtype']]['week'] = $exeCall['week'];
        $arrUserCall[$exeCall['userid']][$exeCall['week'] . $exeCall['agtype']]['dayCall'] = $exeCall['dayCall'];
        $arrUserCall[$exeCall['userid']][$exeCall['week'] . $exeCall['agtype']]['agtype'] = $exeCall['agtype'];
    }
}


$strTotVisit = "SELECT userid, fullname, SUM(TotalCalls) as TotalVisit FROM 
                    (
                    SELECT l.branch, l.fullname, count(userid) as TotalCalls ,userid
                                FROM erpapp.executiveData 
                                LEFT JOIN login l ON l.id = executiveData.userid
                                WHERE dateTime between '" . $firstDay . "' AND '" . $lastDay . "'   
                                    AND l.dept = 'Web-Business'
                                group by WEEK(dateTime), userid
                                order by l.branch,dateTime desc
                    ) totalvis  group by userid";


$result = mysqli_query($con, $strTotVisit);

while ($res = mysqli_fetch_array($result)) {
    $arrTotCalls[] = $res;
}

foreach ($arrUserCall as $key => $arrUser) {
    foreach ($arrTotCalls as $arrTotCall) {

        if (!empty($arrTotCall['userid']) && $arrTotCall['userid'] == $key) {
            $arrUserCall[$key]['total_visit'] = $arrTotCall['TotalVisit'];
        }
    }
}


#####################################################################################################

$strUniCalls = "SELECT userid, clientName,fullname, SUM(unqCalls) as uniqueCalls FROM 
                (            
                        SELECT l.branch, l.fullname, count(distinct userid) as unqCalls,userid,clientName
                        FROM erpapp.executiveData 
                        LEFT JOIN login l ON l.id = executiveData.userid
                        WHERE dateTime between '" . $firstDay . "' AND '" . $lastDay . "'   
                        AND l.dept = 'Web-Business'
                        group by userid, clientName
                        order by l.branch,dateTime desc
                   ) uniquevisit  group by userid ";


$reUnique = mysqli_query($con, $strUniCalls);

while ($resUnq = mysqli_fetch_array($reUnique)) {
    $arrUniqueCalls[] = $resUnq;
}

foreach ($arrUserCall as $key => $arrUser) {
    foreach ($arrUniqueCalls as $arrUnqCall) {

        if (!empty($arrUnqCall['userid']) && $arrUnqCall['userid'] == $key) {
            $arrUserCall[$key]['unique_visit'] = $arrUnqCall['uniqueCalls'];
        }
    }
}

#####Executive Details Calls #################

echo $query = "SELECT l.branch, executiveData.* FROM erpapp.executiveData LEFT join login l ON l.id = executiveData.userid 
where l.dept = 'Web-Business' AND executiveData.dateTime between '" . $firstDay . "' AND '" . $lastDay . "' ORDER BY l.branch,executiveData.dateTime desc";

$result = mysqli_query($con, $query);

$arrAllCalls = array();

while ($res = mysqli_fetch_assoc($result)) {
    $arrAllCalls[] = $res;
}

$adminReport = exportAllBranchReport($arrUserCall, $arrAllCalls);

echo "<br><br>Send Report";
#print_r($arrAllCalls);
#die;

function exportAllBranchReport($arrUserCall, $arrAllCalls) {

    $arrRowData = array();
    $arrWeekI = array();
    $arrWeekII = array();
    $arrWeekIII = array();
    $arrWeekIV = array();

    foreach ($arrUserCall as $arrRow) {
        foreach ($arrRow as $key => $data) {
            if ($key != 'total_visit' && $key != 'unique_visit') {

                $arrRowData[$data['userid']]['username'] = $data['username'];
                $arrRowData[$data['userid']]['branch'] = $data['branch'];
                $arrRowData[$data['userid']]['total_visit'] = $arrRow['total_visit'];
                $arrRowData[$data['userid']]['unique_visit'] = $arrRow['unique_visit'];
                $arrRowData[$data['userid']]['userid'] = $data['userid'];

                if (!empty($data['week']) && $data['week'] == 1) {

                    if (empty($arrRow['1AGENCY']['dayCall'])) {
                        $arrRow['1AGENCY']['dayCall'] = 0;
                    }

                    if (empty($arrRow['1DIRECT']['dayCall'])) {
                        $arrRow['1DIRECT']['dayCall'] = 0;
                    }

                    $arrWeekI['agency'] = '0';
                    $arrWeekI['direct'] = '0';
                    $arrWeekI['total'] = '0';

                    if ($arrRow['1AGENCY']['agtype'] == 'AGENCY') {
                        $arrWeekI['agency'] = $arrRow['1AGENCY']['dayCall'];
                    }

                    if ($arrRow['1DIRECT']['agtype'] == 'DIRECT') {
                        $arrWeekI['direct'] = $arrRow['1DIRECT']['dayCall'];
                    }

                    $arrWeekI['total'] = $arrWeekI['agency'] + $arrWeekI['direct'];
                    $arrRowData[$data['userid']]['week'][1] = $arrWeekI;
                } else {

                    $arrWeekI['agency'] = '0';
                    $arrWeekI['direct'] = '0';
                    $arrWeekI['total'] = '0';

                    $arrRowData[$data['userid']]['week'][1] = $arrWeekI;
                }

                if (!empty($data['week']) && $data['week'] == 2) {

                    $arrWeekII['agency'] = '0';
                    $arrWeekII['direct'] = '0';
                    $arrWeekII['total'] = '0';

                    if (empty($arrRow['2AGENCY']['dayCall'])) {
                        $arrRow['2AGENCY']['dayCall'] = 0;
                    }

                    if (empty($arrRow['2DIRECT']['dayCall'])) {
                        $arrRow['2DIRECT']['dayCall'] = 0;
                    }

                    if ($arrRow['2AGENCY']['agtype'] == 'AGENCY') {
                        $arrWeekII['agency'] = $arrRow['2AGENCY']['dayCall'];
                    }
                    if ($arrRow['2DIRECT']['agtype'] == 'DIRECT') {
                        $arrWeekII['direct'] = $arrRow['2DIRECT']['dayCall'];
                    }

                    $arrWeekII['total'] = $arrWeekII['agency'] + $arrWeekII['direct'];
                    $arrRowData[$data['userid']]['week'][2] = $arrWeekII;
                } else {
                    $arrWeekII['agency'] = '0';
                    ;
                    $arrWeekII['direct'] = '0';
                    ;
                    $arrWeekII['total'] = '0';
                    ;
                }

                if (!empty($data['week']) && $data['week'] == 3) {

                    $arrWeekIII['agency'] = '0';
                    ;
                    $arrWeekIII['direct'] = '0';
                    ;
                    $arrWeekIII['total'] = '0';
                    ;

                    if (empty($arrRow['3AGENCY']['dayCall'])) {
                        $arrRow['3AGENCY']['dayCall'] = 0;
                    }

                    if (empty($arrRow['3DIRECT']['dayCall'])) {
                        $arrRow['3DIRECT']['dayCall'] = 0;
                    }

                    if ($arrRow['3AGENCY']['agtype'] == 'AGENCY') {
                        $arrWeekIII['agency'] = $arrRow['3AGENCY']['dayCall'];
                    }

                    if ($arrRow['3DIRECT']['agtype'] == 'DIRECT') {
                        $arrWeekIII['direct'] = $arrRow['3DIRECT']['dayCall'];
                    }

                    $arrWeekIII['total'] = $arrWeekIII['agency'] + $arrWeekIII['direct'];
                    $arrRowData[$data['userid']]['week'][3] = $arrWeekIII;
                } else {
                    $arrWeekIII['agency'] = '0';
                    ;
                    $arrWeekIII['direct'] = '0';
                    ;
                    $arrWeekIII['total'] = '0';
                    ;
                }

                if (!empty($data['week']) && $data['week'] == 4) {

                    if (empty($arrRow['4AGENCY']['dayCall'])) {
                        $arrRow['4AGENCY']['dayCall'] = 0;
                    }

                    if (empty($arrRow['4DIRECT']['dayCall'])) {
                        $arrRow['4DIRECT']['dayCall'] = 0;
                    }

                    $arrWeekIV['agency'] = '0';
                    ;
                    $arrWeekIV['direct'] = '0';
                    ;
                    $arrWeekIV['total'] = '0';
                    ;

                    if ($arrRow['4AGENCY']['agtype'] == 'AGENCY') {
                        $arrWeekIV['agency'] = $arrRow['4AGENCY']['dayCall'];
                    }

                    if ($arrRow['4DIRECT']['agtype'] == 'DIRECT') {
                        $arrWeekIV['direct'] = $arrRow['4DIRECT']['dayCall'];
                    }

                    $arrWeekIV['total'] = $arrWeekIV['agency'] + $arrWeekIV['direct'];

                    $arrRowData[$data['userid']]['week'][4] = $arrWeekIV;
                } else {
                    $arrWeekIV['agency'] = '0';
                    ;
                    $arrWeekIV['direct'] = '0';
                    ;
                    $arrWeekIV['total'] = '0';
                    ;
                }

                if (!empty($data['week']) && $data['week'] == 5) {

                    if (empty($arrRow['5AGENCY']['dayCall'])) {
                        $arrRow['5AGENCY']['dayCall'] = 0;
                    }

                    if (empty($arrRow['5DIRECT']['dayCall'])) {
                        $arrRow['5DIRECT']['dayCall'] = 0;
                    }

                    $arrWeekV['agency'] = '0';
                    ;
                    $arrWeekV['direct'] = '0';
                    ;
                    $arrWeekV['total'] = '0';
                    ;

                    if ($arrRow['5AGENCY']['agtype'] == 'AGENCY') {
                        $arrWeekV['agency'] = $arrRow['5AGENCY']['dayCall'];
                    }

                    if ($arrRow['5DIRECT']['agtype'] == 'DIRECT') {
                        $arrWeekV['direct'] = $arrRow['5DIRECT']['dayCall'];
                    }

                    $arrWeekV['total'] = $arrWeekV['agency'] + $arrWeekV['direct'];

                    $arrRowData[$data['userid']]['week'][5] = $arrWeekV;
                } else {
                    $arrWeekV['agency'] = '0';
                    ;
                    $arrWeekV['direct'] = '0';
                    ;
                    $arrWeekV['total'] = '0';
                    ;
                }
            }
        }
    }

    $arrweek = array(1, 2, 3, 4, 5);

    foreach ($arrRowData as $rowData) {
        foreach ($rowData as $key => $data) {
            if (is_array($rowData[$key]) && !array_key_exists($arrweek[0], $rowData['week'])) {

                $weekI['agency'] = '0';
                $weekI['direct'] = '0';
                $weekI['total'] = '0';

                $arrRowData[$rowData['userid']]['week'][1] = $weekI;
            }

            if (is_array($rowData[$key]) && !array_key_exists($arrweek[1], $rowData['week'])) {
                $weekII['agency'] = '0';
                $weekII['direct'] = '0';
                $weekII['total'] = '0';

                $arrRowData[$rowData['userid']]['week'][2] = $weekII;
            }
            if (is_array($rowData[$key]) && !array_key_exists($arrweek[2], $rowData['week'])) {
                $weekIII['agency'] = '0';
                $weekIII['direct'] = '0';
                $weekIII['total'] = '0';

                $arrRowData[$rowData['userid']]['week'][3] = $weekIII;
            }

            if (is_array($rowData[$key]) && !array_key_exists($arrweek[3], $rowData['week'])) {

                $weekIV['agency'] = '0';
                $weekIV['direct'] = '0';
                $weekIV['total'] = '0';

                $arrRowData[$rowData['userid']]['week'][4] = $weekIV;
            }

            if (is_array($rowData[$key]) && !array_key_exists($arrweek[4], $rowData['week'])) {

                $weekV['agency'] = '0';
                $weekV['direct'] = '0';
                $weekV['total'] = '0';

                $arrRowData[$rowData['userid']]['week'][5] = $weekV;
            }

            ksort($arrRowData[$rowData['userid']]['week']);
        }
    }
    
    #echo "<pre>XXX"; print_r($arrRowData); die;


    echo date('H:i:s'), " Create new PHPExcel object", EOL;
    $objPHPExcel = new PHPExcel();

    // Set document properties
    echo date('H:i:s'), " Set document properties", EOL;
    $objPHPExcel->getProperties()->setCreator("Gurunath")
            ->setLastModifiedBy("Gurunath")
            ->setTitle("Report of : Gurunath")
            ->setSubject("Dated")
            ->setDescription("Automated Report by Gurunath")
            ->setKeywords("CRM Web Sales REPORT")
            ->setCategory("CRM Web Sales Report");


    $styleArray = array(
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN
            )
        )
    );

    $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getActiveSheet()->setTitle('Call Summary');
    $objPHPExcel->getActiveSheet()->getStyle('A1:S4')->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()->setARGB('FFFFCC99');
    $objPHPExcel->getActiveSheet()->getStyle('A1:S4')->getFont()->setBold(true);

    $xlsdata = array();
    $cnt = 0;

    $xlsdata[0] = array('Space Marketing - Market Visit Web Sales Report for the month of ' . date('M Y'));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:C1');

    $xlsdata[1] = array('Report updated upto ' . date('l jS M Y'));
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells('A2:C2');

    $objPHPExcel->getActiveSheet()->getRowDimension(2)->setRowHeight(20);

    $xlsdata[2] = array('', '', '', '', 'Week 1', '', '', 'Week 2', '', '', 'Week 3', '', '', 'Week 4', '', '', 'Week 5');

    $objPHPExcel->getActiveSheet()->getRowDimension(2)->setRowHeight(20);

    $xlsdata[3] = array("Web Sales Executive", "Location", "Total Visit in a Day", "Unique Visit in a Day", "Agency", "Direct", "Week I Total", "Agency", "Direct", "Week II Total", "Agency", "Direct", "Week III Total", "Agency", "Direct", "Week IV Total", "Agency", "Direct", "Week V Total");

    $objPHPExcel->setActiveSheetIndex(0)->mergeCells('E3:G3');
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells('H3:J3');
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells('K3:M3');
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells('N3:P3');
    $objPHPExcel->setActiveSheetIndex(0)->mergeCells('Q3:S3');

    $objPHPExcel->getActiveSheet()->fromArray($xlsdata, NULL, 'A' . ($cnt + 1));


    foreach (range('A', 'S') as $columnID) {
        $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
    }

    $cnt = 4;
    foreach ($arrRowData as $rowData) {
        $xlsdata[$cnt][] = $rowData['username'];
        $xlsdata[$cnt][] = $rowData['branch'];
        $xlsdata[$cnt][] = $rowData['total_visit'];
        $xlsdata[$cnt][] = $rowData['unique_visit'];

        foreach ($arrRowData[$rowData['userid']]['week'] as $wkey => $wvalue) {

            $xlsdata[$cnt][] = $wvalue['agency'];
            $xlsdata[$cnt][] = $wvalue['direct'];
            $xlsdata[$cnt][] = $wvalue['total'];

            $objPHPExcel->getActiveSheet()->fromArray($xlsdata[$cnt], NULL, 'A' . ($cnt + 1));
        }

        $cnt++;
    }
    $objPHPExcel->getActiveSheet()->getStyle('A1:S' . $cnt)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->fromArray($xlsdata[0], NULL, 'A1');

    $objPHPExcel->createSheet();
    $objPHPExcel->setActiveSheetIndex(1);
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'More data');
    $objPHPExcel->getActiveSheet()->setTitle('Call Details');
    $objPHPExcel->getActiveSheet()->getStyle('A1:S1')->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()->setARGB('FFFFCC99');
    $objPHPExcel->getActiveSheet()->getStyle('A1:S1')->getFont()->setBold(true);

    $xlsdata = array();
    $cnt = 0;

    $xlsdata[0] = array('Branch', 'Executive', 'clientName', 'personName', 'summary', 'escalate', 'region', 'location', 'address1', 'address2', 'dateTime', 'userid', 'clientCode', 'contact_no', 'followup_date', 'reason_id', 'email_id', 'mobile_no', 'agtype');

    $objPHPExcel->getActiveSheet()->fromArray($xlsdata[0], NULL, 'A1');
    $objPHPExcel->getActiveSheet()->getRowDimension(10)->setRowHeight(20);

    foreach (range('A', 'S') as $columnID) {
        $objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
    }


    $cnt = 1;

    foreach ($arrAllCalls as $exeData) {

        $xlsdata[$cnt][] = $exeData['branch'];
        $xlsdata[$cnt][] = $exeData['username'];
        $xlsdata[$cnt][] = $exeData['clientName'];
        $xlsdata[$cnt][] = $exeData['personName'];
        $xlsdata[$cnt][] = $exeData['summary'];
        $xlsdata[$cnt][] = $exeData['escalate'];
        $xlsdata[$cnt][] = $exeData['region'];
        $xlsdata[$cnt][] = $exeData['location'];
        $xlsdata[$cnt][] = $exeData['address1'];
        $xlsdata[$cnt][] = $exeData['address2'];
        $xlsdata[$cnt][] = $exeData['dateTime'];
        $xlsdata[$cnt][] = $exeData['userid'];
        $xlsdata[$cnt][] = $exeData['clientCode'];
        $xlsdata[$cnt][] = $exeData['contact_no'];
        $xlsdata[$cnt][] = $exeData['followup_date'];
        $xlsdata[$cnt][] = $exeData['reason_id'];
        $xlsdata[$cnt][] = $exeData['email_id'];
        $xlsdata[$cnt][] = $exeData['mobile_no'];
        $xlsdata[$cnt][] = $exeData['agtype'];

        $objPHPExcel->getActiveSheet()->fromArray($xlsdata[$cnt], NULL, 'A' . ($cnt + 1));

        $cnt++;
    }

    $objPHPExcel->getActiveSheet()->getStyle('A1:S' . $cnt)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->fromArray($xlsdata[0], NULL, 'A1');

    $objPHPExcel->setActiveSheetIndex(0);

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $objWriter->save('/usr1/htdocs/erp/bs/crons/market_visit/web_market_visit.xls');

    $filename = 'web_market_visit.xls';

    $sendtype = '';

    # $fdate = date('Y-m-d', strtotime("Last Sunday -6 days"));
    $fdate = date('Y-m-d');

    $tdate = date('Y-m-d', strtotime("Last Sunday"));

    $attachmentFile = $filename;
    $fromName = "CRM WEB Sales Week wise Call Report";
    $fromMail = "From: CRM BS-Web Sales <webtech@bsmail.in>";
    #$fromMail = "Weekly Report<socials@scmbs.com>";
    #$subject  = "CRM Market Visit Call - " . date('l jS M Y', strtotime('Last Sunday'));
    $subject = "CRM WEB Sales Market Visit Call Report Till  " . date('l jS M Y');
    $message = '<p>Dear Sir,</p>';
    # $message  .= '<p>PFA <strong>CRM Market Visit Call Data & Summary </strong> report till '.date('dS F Y',strtotime($fdate)).' to '.date('dS F Y',strtotime($tdate)).'</p>';
    $message .= '<p>PFA <strong>CRM WEB Sales Market Visit Call Data & Summary </strong> report till ' . date('dS F Y', strtotime($fdate)) . ' </p>';

    $message .= '<strong>Note : </strong>This is an auto generated report, there could be stray cases of technical glitches wherein the data may show up as blank. Efforts have been taken to minimize this';
    $message .= '<p>Regards,</p>';
    $message .= '<p>CRM WEB Sales </p>';

    libmail_attach($attachmentFile, "shailendra.kalelkar@bsmail.in", $fromMail, $fromName, $subject, $message, $sendtype);
   // libmail_attach($attachmentFile, "gurunath.auti@bsmail.in", $fromMail, $fromName, $subject, $message, $sendtype);


    #echo "<pre>dddd";
    #print_r($arrRowData); die; 

    mylog("Advt Trade Call All Branch Weekwise Report stored in $name.csv", 'Notice');

    return $name;
}

function mylog($msg, $type = 'notice') {

    $line = "\n" . date('Y-m-d H:i:s') . "\t" . $type . "\t" . $msg;
    file_put_contents('pending_log.txt', $line, FILE_APPEND);
}

function libmail_attach($filename, $mailto, $from_mail, $from_name, $subject, $message, $sendtype = "internal") {
    //echo "hiiii";
    //echo $sendtype;die;
    $file = '/usr1/htdocs/erp/bs/crons/market_visit/' . $filename;
    $m = new Mail(); // create the mail
    $m->From($from_mail);
    $m->To($mailto);
    $m->Subject($subject);
    $m->Body($message);

    //$m->Bcc($bcc);
    $m->Cc("gurunath.auti@bsmail.in");
    $m->Priority(4);
    //	attach a file of type image/gif to be displayed in the message if possible
    $m->Attach($file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "inline");
    $m->Send(); // send the mail
    echo "Mail was sent:";
    //echo $m->Get(); // show the mail source
}


?>