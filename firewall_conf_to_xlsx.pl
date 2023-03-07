#!/usr/bin/perl

use strict;
use warnings;
use v5.28;
use Excel::Writer::XLSX;
use Encode;
use POSIX qw/strftime/;
binmode(STDOUT, ":encoding(cp936)");

# 获取参数并判断文件是否为文本文件
my $file_path;
if( scalar(@ARGV) != 1 ){
    say "This script accepts only one parameter.\n  Usage: $0 firewall-config-file.txt";
    exit(1);
}else{
    $file_path = $ARGV[0];
    unless( -e $file_path and -T $file_path ){
        say "File $file_path does not exist or is not a text file.";
        exit(2);
    }
}

# 定义输出的 xlsx 文件名称
my $xlsx = $file_path ;
my $date = POSIX::strftime("%Y%m%d_%H%M%S", localtime);

# 如果源文件名包含后缀，则新的xlsx文件名为去掉后缀并加入时间
# 否则直接加入时间
if($xlsx =~ m|\.[^\./]+$|){
    $xlsx =~ s|\.[^\./]+$|_$date.xlsx|;
}else{
    $xlsx = $xlsx . "_$date.xlsx";
}

say("Input: $file_path , output: $xlsx","\n","Processing...");

# 开始处理
open(my $config_file, '<:encoding(cp936)', $file_path)
    or die("Error: $!");
my @config_Text = <$config_file>;
close($config_file);

# 定义数据结构（可省略）
my %config_hash = (
    interface             => undef, ##   ! interface
    address               => undef, ##   ! address
    address_group         => undef, ##   ! address-group
    service               => undef, ##   ! service
    service_group         => undef, ##   ! service-group
    schedule              => undef, ##   ! schedule
    firewall_policy_group => undef, ##   ! firewall policy group ipv4
    firewall_policy       => undef  ##   ! firewall policy
);

# 定义当前行号和下一行号变量
my ($curr_line, $next_line) = ();
# 定义当前行和下一行内容变量
my ($curr_text, $next_text) = ();
# 各段数据提取开关
my $interface_sw = "off";
my $address_sw = "off";
my $address_group_sw = "off";
my $service_sw = "off";
my $service_group_sw = "off";
my $schedule_sw = "off";
my $firewall_policy_group_sw = "off";
my $firewall_policy_sw = "off";

# 定义正则表达式分组数据缓存变量
my ($G1, $G2, $G3, $G4, $G5, $G6, $G61, $G7, $G8) = ();

## 为避免变量 $next_line 越界，所以减2
my $total_line = scalar(@config_Text) - 2;

# 遍历文件的每一行
for ($curr_line = 0; $curr_line < $total_line; $curr_line++) {

    # 获取当前行内容及下一行行号及内容，并将内容的行尾空格及回车符删除
    $next_line = $curr_line + 1;
    $curr_text = $config_Text[$curr_line];
    $next_text = $config_Text[$next_line];
    $curr_text =~ s/\s+$//;
    $next_text =~ s/\s+$//;
    chomp($curr_text);
    chomp($next_text);

    # 如果当前行是感叹号!，则判断下行开始的数据类型
    if ($curr_text eq "!") {
        # 数据提取-开关
        # 如果下一行是 interface 开头，则开关打开，否则关闭
        if ($next_text =~ m/^interface (.+)/) {
            $G1 = $1;
            $config_hash{interface}->{$G1} = undef;
            $interface_sw = "on";
            $curr_line += 1;
        }
        else {
            $interface_sw = "off";
        }
        # 数据提取-开关
        # 如果下一行是 address 开头，则开关打开，否则关闭
        if ($next_text =~ m/^address (.+)/) {
            $address_sw = "on";
        }
        else {
            $address_sw = "off";
        }
        # 数据提取-开关
        # 如果下一行是 address-group 开头，则开关打开，否则关闭
        if ($next_text =~ m/^address-group (.+)/) {
            $address_group_sw = "on";
        }
        else {
            $address_group_sw = "off";
        }
        # 数据提取-开关
        # 如果下一行是 service 开头，则开关打开，否则关闭
        if ($next_text =~ m/^service (.+)/) {
            $service_sw = "on";
        }
        else {
            $service_sw = "off";
        }
        # 数据提取-开关
        # 如果下一行是 service-group 开头，则开关打开，否则关闭
        if ($next_text =~ m/^service-group (.+)/) {
            $service_group_sw = "on";
        }
        else {
            $service_group_sw = "off";
        }
        # 数据提取-开关
        # 如果下一行是 schedule 开头，则开关打开，否则关闭
        if ($next_text =~ m/^schedule (recurring|onetime) (.+)/) {
            $schedule_sw = "on";
        }
        else {
            $schedule_sw = "off";
        }
        # 数据提取-开关
        # 如果下一行是 firewall policy group 开头，则开关打开，否则关闭
        if ($next_text =~ m/^firewall policy group (.+)/) {
            $firewall_policy_group_sw = "on";
        }
        else {
            $firewall_policy_group_sw = "off";
        }
        # 数据提取-开关
        # 如果下一行是 firewall policy 开头，则开关打开，否则关闭
        if ($next_text =~ m/^firewall policy (?!group)/) {
            $firewall_policy_sw = "on";
        }
        else {
            $firewall_policy_sw = "off";
        }
    }
    #  interface 数据提取
    elsif ($interface_sw eq "on") {
        if ($curr_text =~ /^ +(float ip address) (.+)/) {
            $config_hash{interface}->{$G1}->{$1} = $2;
        }
        elsif ($curr_text =~ /^ +(ip address) (.+)/) {
            $config_hash{interface}->{$G1}->{$1} = $2;
        }
        elsif ($curr_text =~ /^ +([^ ]+) (.+)/) {
            $config_hash{interface}->{$G1}->{$1} = $2;
        }
        else {
            say "- Interface match nothing -";
        }
    }
    #  address-object 数据提取
    elsif ($address_sw eq "on") {
        if ($curr_text =~ /^address (.+)/) {
            $G2 = $1;
        }
        elsif ($curr_text =~ /^description (.+)/) {
            push @{$config_hash{address}->{$G2}->{description}}, $2;
        }
        elsif ($curr_text =~ /^ +([^ ]+) (.+)/) {
            push @{$config_hash{address}->{$G2}->{$1}}, $2;
        }
        else {
            say "- address-object match nothing -";
        }
    }
    #  address-group 数据提取
    elsif ($address_group_sw eq "on") {
        if ($curr_text =~ /^address-group (.+)/) {
            $G3 = $1;
        }
        elsif ($curr_text =~ /^ +([^ ]+) (.+)/) {
            push @{$config_hash{address_group}->{$G3}->{$1}}, $2;
        }
        else {
            say "- address-group match nothing -";
        }
    }
    #  service 数据提取
    elsif ($service_sw eq "on") {
        if ($curr_text =~ /^service (.+)/) {
            $G4 = $1;
        }
        elsif ($curr_text =~ /^ +(description) (.+)/) {
            $config_hash{service}->{$G4}->{$1} = $2;
        }
        elsif ($curr_text =~ /^ +(\S+) +(.+)/) {
            push @{$config_hash{service}->{$G4}->{$1}}, $2;
        }
        else {
            say "- service match nothing -";
        }
    }
    #  service-group 数据提取
    elsif ($service_group_sw eq "on") {
        if ($curr_text =~ /^service-group (.+)/) {
            $G5 = $1;
        }
        elsif ($curr_text =~ /^ +(\S+) +(.+)/) {
            push @{$config_hash{service_group}->{$G5}->{$1}}, $2;
        }
        else {
            say "- service-group match nothing -";
        }
    }
    #  schedule 数据提取
    elsif ($schedule_sw eq "on") {
        if ($curr_text =~ /^schedule recurring (.+)/) {
            $G6 = "recurring";
            $G61 = $1;
        }
        elsif ($curr_text =~ /^schedule onetime (.+)/) {
            $G6 = "onetime";
            $G61 = $1;
        }
        elsif ($curr_text =~ /^ +(\S+) +(.+)/) {
            $config_hash{schedule}->{$G6}->{$G61} = $2;
        }
        else {
            say "- schedule match nothing -";
        }
    }
    #  firewall policy group 数据提取
    elsif ($firewall_policy_group_sw eq "on") {
        if ($curr_text =~ /^firewall policy group ipv4 (.+)/) {
            $G7 = $1;
            $config_hash{firewall_policy_group}->{$G7} = $1;
        }
        else {
            say "- firewall policy group match nothing -";
        }
    }
    #  firewall policy 数据提取
    elsif ($firewall_policy_sw eq "on") {
        if ($curr_text =~ /^firewall policy (?!group)(.+)/) {
            $G8 = $1;
        }
        elsif ($curr_text =~ /^ +(action|name|user|app|timerange|description|firewall-policy-group) +(.+)/) {
            $config_hash{firewall_policy}->{$G8}->{$1} = $2;
        }
        elsif ($curr_text =~ /^ +enable$/) {
            $config_hash{firewall_policy}->{$G8}->{enable} = "true";
        }
        elsif ($curr_text =~ /^ +flowstat$/) {
            $config_hash{firewall_policy}->{$G8}->{flowstat} = "true";
        }
        elsif ($curr_text =~ /^ +(\S+) +(.+)/) {
            push @{$config_hash{firewall_policy}->{$G8}->{$1}}, $2;
        }
        else {
            say "- firewall policy match nothing -";
        }
    }
    else {
        # say "ignore line number: $curr_line --> $curr_text";
    }
}

# 创建 xlsx 文件
my $workbook = Excel::Writer::XLSX->new($xlsx);

# 增加sheet页
my $sheet_name = decode("gbk", "防火墙策略清单");
my $worksheet = $workbook->add_worksheet($sheet_name);

# 冻结第一行(1,0) ,如果冻结第一列，则(0,1)，第一行第一列同时冻结，则(1,1)
$worksheet->freeze_panes(1,0);

# 标题格式
my $format_title = $workbook->add_format();
$format_title->set_bold(1);  # 粗体
$format_title->set_size(10); # 字体大小
$format_title->set_bg_color('silver');  # 背景颜色
$format_title->set_align('top');  # 对齐模式
$format_title->set_border(1);  # 边框

# 数据格式
my $format_data = $workbook->add_format();
$format_data->set_size(10);
$format_data->set_align('vcenter');

# 表头,从A列开始，A 的 ascii值 65
my $col_ascii = 65;
$worksheet->write( chr($col_ascii)   . "1", decode("gbk", "策略序号(policyID)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "策略名称(name)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "是否启用(enable)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "策略行为(action)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "源端地址名称(src-addr-name)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "源端地址(src-address)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "源端地址组名称(src-group-name)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "源端地址组地址(src-group-address)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "源端区域(src-zone)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "目标区域(dst-zone)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "目标地址名称(dst-addr)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "目标地址(dst-address)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "目标地址组名称(dst-group-addr)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "目标地址组地址(dst-group-address)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "目标端口名称(service)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "目标端口(service_port)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "目标端口描述(service_desc)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "生效时间(timerange)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "防火墙策略组(firewall-policy-group)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "策略描述(description)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "用户(user)"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "app"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "flowstat"), $format_title);
$worksheet->write( chr(++$col_ascii) . "1", decode("gbk", "日志(log)"), $format_title);

my $Line = 1;
for my $i (sort {$a <=> $b} keys %{$config_hash{firewall_policy}}) {
    # 策略名称
    my $name = $config_hash{firewall_policy}->{$i}->{name} // "";
    # 是否启用
    my $enable = $config_hash{firewall_policy}->{$i}->{enable} // "false";
    # 动作
    my $action = $config_hash{firewall_policy}->{$i}->{action} // "";
    # 获取策略中的源端地址名称，然后在地址段address中将名称解析为具体的地址
    my $src_addr_name = "";
    my $src_addr_detail = "";
    my $src_num = 1;
    for my $a1 (@{$config_hash{firewall_policy}->{$i}->{"src-addr"}}) {
        # 判断如果是源地址名称中存在地址组，则跳过，地址组后面专门处理
        if(defined($config_hash{address_group}->{$a1})){
            next;
        }
        $src_addr_name .= $src_num . ":" . $a1 . "\n";
        for my $address_keys (keys %{$config_hash{address}->{$a1}}) {
            # 忽略详细地址的描述信息
            if($address_keys ne "description"){
                for my $a2 (@{$config_hash{address}->{$a1}->{$address_keys}}) {
                    $src_addr_detail .= $src_num . ":" .$a2 . "\n";
                }
            }
        }
        $src_num++;
    }
    # 获取源地址组
    my $src_addr_group = "" ; # 默认为空，没有地址组时也不会出现标量值未初始化的告警
    my $src_addr_group_detail = ""; # 默认为空，没有地址组时也不会出现标量值未初始化的告警
    my $src_addr_grp_num = 1;
    for my $a7 (@{$config_hash{firewall_policy}->{$i}->{"src-addr"}}){
        # 判断如果源地址是地址组则继续处理
        if(defined($config_hash{address_group}->{$a7})){
            for my $a8 (@{$config_hash{address_group}->{$a7}->{"address-object"}}){
                # 拼接地址组中每个地址对象的名称
                $src_addr_group_detail .= $src_addr_grp_num . ":" . $a8 . ":\n";
                for my $address_keys (keys %{$config_hash{address}->{$a8}}) {
                    # 忽略详细地址的描述信息
                    if($address_keys ne "description"){
                        for my $a9 (@{$config_hash{address}->{$a8}->{$address_keys}}) {
                            $src_addr_group_detail .= " "x2 . $a9 . "\n";
                        }
                    }
                }
            }
            $src_addr_group .= $src_addr_grp_num . ":" . $a7 . "\n";
            $src_addr_grp_num++;
        }
    }
    # 获取源端区域
    my $src_zone = "";
    my $src_inter_detail = "";
    for my $b1 (@{$config_hash{firewall_policy}->{$i}->{"src-zone"}}) {
        for my $keys_1 (sort { $a cmp $b } keys %{$config_hash{interface}->{$b1}}){
            $src_inter_detail .= " "x4 . $keys_1 . ":". $config_hash{interface}->{$b1}->{$keys_1} . "\n";
        }
        $src_zone .= $b1 . ":\n" . "$src_inter_detail" ."\n";
        chomp($src_zone);
    }
    # 获取目标端区域
    my $dst_zone = "";
    my $dst_inter_detail = "" ;
    for my $c1 (@{$config_hash{firewall_policy}->{$i}->{"dst-zone"}}) {
        for my $keys_2 (sort { $a cmp $b } keys %{$config_hash{interface}->{$c1}}){
            $dst_inter_detail .= " "x4 . $keys_2 . ":". $config_hash{interface}->{$c1}->{$keys_2} . "\n";
        }
        $dst_zone .= $c1 . ":\n" . "$dst_inter_detail" ."\n";
        chomp($dst_zone);
    }
    # 获取策略中的源端地址名称，然后在地址段address中将名称解析为具体的地址
    my $dst_addr_name = "";
    my $dst_addr_detail = "";
    my $dst_num = 1;
    for my $d1 (@{$config_hash{firewall_policy}->{$i}->{"dst-addr"}}) {
        # 判断如果是目标地址名称中存在地址组，则跳过，地址组后面专门处理
        if(defined($config_hash{address_group}->{$d1})){
            next;
        }
        $dst_addr_name .= $dst_num . ":" . $d1 . "\n";
        for my $address_keys (keys %{$config_hash{address}->{$d1}}) {
            for my $d2 (@{$config_hash{address}->{$d1}->{$address_keys}}) {
                $dst_addr_detail .= $dst_num . ":" . $d2 . "\n";
            }
        }
        $dst_num++;
    }
    # 获取目标地址组
    my $dst_addr_group = "" ; # 默认为空，没有地址组时也不会出现标量值未初始化的告警
    my $dst_addr_group_detail = ""; # 默认为空，没有地址组时也不会出现标量值未初始化的告警
    my $dst_addr_grp_num = 1;
    for my $d7 (@{$config_hash{firewall_policy}->{$i}->{"dst-addr"}}){
        if(defined($config_hash{address_group}->{$d7})){
            for my $d8 (@{$config_hash{address_group}->{$d7}->{"address-object"}}){
                # 拼接地址组中每个地址对象的名称
                $dst_addr_group_detail .= $dst_addr_grp_num . ":" . $d8 . ":\n";
                for my $dst_keys (keys %{$config_hash{address}->{$d8}}) {
                    # 忽略详细地址的描述信息
                    if($dst_keys ne "description"){
                        for my $d9 (@{$config_hash{address}->{$d8}->{$dst_keys}}) {
                            $dst_addr_group_detail .= " "x2 . $d9 . "\n";
                        }
                    }
                }
            }
            $dst_addr_group .= $dst_addr_grp_num . ":" . $d7 . "\n";
            $dst_addr_grp_num++;
        }
    }
    # 获取策略中的目标名称，然后在地址段 service 中将名称解析为具体的端口
    my $service = "";
    my $service_port = "";
    my $service_desc = "";
    my $srv_num = 1;
    for my $e1 (@{$config_hash{firewall_policy}->{$i}->{service}}) {
        $service .= $srv_num . ":" . $e1 . "\n";
        # 根据端口策略名称，获取协议
        for my $p_keys (keys %{$config_hash{service}->{$e1}}) {
            # 如果端口数据中有描述信息（description），则提取描述信息；
            if ($p_keys eq "description") {
                $service_desc .= $config_hash{service}->{$e1}->{$p_keys} . "\n";
            }
            else {
                # 将端口数据提取出来，在前面加上端口协议
                for my $port (@{$config_hash{service}->{$e1}->{$p_keys}}) {
                    $service_port .= $srv_num . ":" . $p_keys . ":" . $port . "\n";
                }
            }
        }
        $srv_num++;
    }
    # 生效时间
    my $timerange;
    if ($config_hash{firewall_policy}->{$i}->{timerange} eq "always") {
        # 如果是永久生效为 always
        $timerange = "always";
    }
    else {
        # 如果不是always，则获取具体数据
        my $timerKey = $config_hash{firewall_policy}->{$i}->{timerange};
        # 判断是一次性生效 onetime 还是经常性 recurring
        if (defined($config_hash{schedule}->{onetime}->{$timerKey})) {
            $timerange = "onetime:" . $config_hash{schedule}->{onetime}->{$timerKey};
        }
        elsif (defined($config_hash{schedule}->{recurring}->{$timerKey})) {
            $timerange = "recurring:" . $config_hash{schedule}->{recurring}->{$timerKey};
        }
        else {
            $timerange = "null";
        }
    }
    # 防火墙策略组
    my $firewall_policy_group = $config_hash{firewall_policy}->{$i}->{"firewall-policy-group"} // "";
    # 策略描述
    my $description = $config_hash{firewall_policy}->{$i}->{description} // "";
    # 用户
    my $user = $config_hash{firewall_policy}->{$i}->{user} // "";
    # 应用
    my $app = $config_hash{firewall_policy}->{$i}->{app} // "";
    # flowstat
    my $flowstat = $config_hash{firewall_policy}->{$i}->{flowstat} // "false";
    # 日志(log)
    my $log_status;
    if(defined($config_hash{firewall_policy}->{$i}->{log})) {
        for my $f1 (@{$config_hash{firewall_policy}->{$i}->{log}}) {
            $log_status .= $f1 . "\n";
        }
    }else{
        $log_status = "";
    }
    # 删除变量末尾的换行符
    chomp($src_addr_name);
    chomp($src_addr_detail);
    chomp($src_addr_group);
    chomp($src_addr_group_detail);
    chomp($src_zone);
    chomp($dst_zone);
    chomp($dst_addr_name);
    chomp($dst_addr_detail);
    chomp($dst_addr_group);
    chomp($dst_addr_group_detail);
    chomp($service);
    chomp($service_port);
    chomp($service_desc);
    chomp($log_status);

    # 向excel文件写入数据
    $Line++;
    my $data_ascii = 65;  # 必须与表头的 $col_ascii 初始值一致，否则数据与表头不一致
    $worksheet->write( chr($data_ascii)   . $Line, "$i", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$name", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$enable", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$action", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$src_addr_name", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$src_addr_detail", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$src_addr_group", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$src_addr_group_detail", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$src_zone", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$dst_zone", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$dst_addr_name", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$dst_addr_detail", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$dst_addr_group", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$dst_addr_group_detail", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$service", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$service_port", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$service_desc", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$timerange", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$firewall_policy_group", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$description", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$user", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$app", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$flowstat", $format_data);
    $worksheet->write( chr(++$data_ascii) . $Line, "$log_status", $format_data);
}

# 关闭xlsx文件写入
$workbook->close();
print("Done");
