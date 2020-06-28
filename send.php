<?php

require "vendor/autoload.php";


if (isset($_FILES['addresser']) && $_FILES['addresser']['size']){
    $addressers = readExcel($_FILES['addresser']['tmp_name']);
}
if (isset($_FILES['addressee']) && $_FILES['addressee']['size']){
    $addressees = readExcel($_FILES['addressee']['tmp_name']);
}

$subject = $_POST['subject'];
$body = $_POST['body'];

echo "<pre>";
foreach ($addressees as $addressee){
    $currentAddresser = current($addressers);
    while ($currentAddresser){
        echo "发送邮件 {$currentAddresser[0]} 至 {$addressee[0]} \n";
        try {
            sendMail($currentAddresser[0], $currentAddresser[1], $addressee[0], $subject, $body);
            echo "发送成功\n";
            break;
        } catch (Exception $e){
            // 记录
            echo "发送失败 messaage：{$e->getMessage()} \n";


            $currentAddresser = next($addressers);
            if (!$currentAddresser){
                echo "无发件人\n";
                break 2;
            }
        }
    }

    echo "\n";
}
echo "</pre>";

function readExcel($filename){
    $addresserExcel = \PhpOffice\PhpSpreadsheet\IOFactory::load($filename);
    $sheet = $addresserExcel->getActiveSheet();

    $res = [];

    foreach ($sheet->getRowIterator(2) as $row) {
        $rowIndex = $row->getRowIndex();
        $res[$rowIndex] = [];
        foreach ($row->getCellIterator() as $cell) {
            $res[$rowIndex][] = $cell->getFormattedValue();
        }
    }

    return $res;
}

function getMailHost($address){
    return substr($address, strpos($address, '@') + 1);
}

function sendMail($from, $password, $to, $subject, $body){
    $mailhostMap = [
        '163.com' => [
            "smtp_host" => 'smtp.163.com',
            "smtp_port" => 25,
        ],
    ];

    $host = getMailHost($from);
    $smtpHost = $mailhostMap[$host]['smtp_host'] ?? '';
    $smtpPort = $mailhostMap[$host]['smtp_port'] ?? 25;

    try {
        $mail = new \PHPMailer\PHPMailer\PHPMailer(true);                              // Passing `true` enables exceptions

        //服务器配置
        $mail->CharSet ="UTF-8";                     //设定邮件编码
        $mail->SMTPDebug = 0;                        // 调试模式输出
        $mail->isSMTP();                             // 使用SMTP
        $mail->Host = $smtpHost;                // SMTP服务器
        $mail->SMTPAuth = true;                      // 允许 SMTP 认证
        $mail->Username = $from;                // SMTP 用户名  即邮箱的用户名
        $mail->Password = $password;             // SMTP 密码  部分邮箱是授权码(例如163邮箱)
        //$mail->SMTPSecure = 'ssl';                    // 允许 TLS 或者ssl协议
        $mail->Port = $smtpPort;                            // 服务器端口 25 或者465 具体要看邮箱服务器支持

        $mail->setFrom($from);  //发件人
        $mail->addAddress($to);  // 收件人
        $mail->addReplyTo($to); //回复的时候回复给哪个邮箱 建议和发件人一致
        //$mail->addCC('cc@example.com');                    //抄送
        //$mail->addBCC('bcc@example.com');                    //密送

        //发送附件
        // $mail->addAttachment('../xy.zip');         // 添加附件
        // $mail->addAttachment('../thumb-1.jpg', 'new.jpg');    // 发送附件并且重命名

        //Content
        $mail->Subject = $subject;
        $mail->Body    = $body;

        $mail->send();
        return true;
    } catch (Exception $e) {
        throw new Exception($mail->ErrorInfo);
    }
}