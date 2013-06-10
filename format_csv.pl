#!/usr/bin/perl

use strict;
use warnings;
use Text::CSV;
use Data::Dumper;
use Spreadsheet::WriteExcel;
use Spreadsheet::WriteExcel::Utility qw( xl_range_formula );

open I, "regions_merged_uterus16s.csv" or die "Can't open the csv file: $!\n";

my $csv = Text::CSV->new();
my @desig;#designation
my %desig_data;#Data in the csv
my $count = 0;
my $messed_data = 0;

while (<I>){
    if ($csv->parse($_)){
	for my $f ($csv->fields()){
	    #if (!defined($f) || $f=~//){next}
	    if ($. == 1){
		$desig[$count]=$f;
	    }
	    else{
		push (@{$desig_data{$desig[$count]}}, $f);
	    }
	    $count++;
	}
	$count=0;
    }
}
my $size = @desig;
#print Dumper(%desig_data);
#print "size of array is $size\n";
#for my $k(@desig){
#    print $k."\n";
#}
my %consolidated;
my @types;
my $curr_d;
for my $n (0..35){
    for my $k (@desig){
	if (!defined($k) || $k=~/sample|MID/){
	    ${$consolidated{$k}}[$n]=${$desig_data{$k}}[$n];
	    push @types, $k unless $n !=0;
	    next;
	}
	if ($k =~ /norank/){
	    $curr_d = $k;
	}
	if ($n==0){
	    my @a = split(/ |_/,$k);
	    push @types, $a[0];
	}

	if ($k!~/norank|^domain|^phylum|^genus|^class|^family|^order/){
	    ${$consolidated{$curr_d}}[$n]+=${$desig_data{$k}}[$n];# or die "n is $n k is $k curr_d is $curr_d\n";
#	    print "Successful consolidation: ${$consolidated{$curr_d}}[$n] and original data: ${$desig_data{$k}}[$n]\n";
	}
	else{
	    $curr_d = $k;
	    ${$consolidated{$curr_d}}[$n]=0;
	    if (${$consolidated{$curr_d}}[$n]!=0){print "Failed to set to zero";}
	    my $var = ${$desig_data{$k}}[$n];
	    if (!defined ${$desig_data{$k}}[$n]){print "It isn't defined!!!!\n";}
	    ${$consolidated{$curr_d}}[$n]+=$var;# or die "else loop, k is $k n is $n, desig val is $var error: $!\n";
#	    print "successful else consolidated: ${$consolidated{$curr_d}}[$n] original data: $var k $k and n is $n\n";
	}
    }
}
my $wb_genus = Spreadsheet::WriteExcel->new('genus.xls');
my $wb_phylum = Spreadsheet::WriteExcel->new('phylum.xls');
my $wb_class = Spreadsheet::WriteExcel->new('class.xls');
my $wb_family = Spreadsheet::WriteExcel->new('family.xls');
my $wb_order = Spreadsheet::WriteExcel->new('order.xls');
open GENUS, ">genus_uterus.tsv" or die "Can't open the genus_uterus.tsv file: $!\n";
open PHYL, ">phylum_uterus.tsv" or die "Can't open the genus_uterus.tsv file: $!\n";
open CLASS, ">class_uterus.tsv" or die "Can't open the genus_uterus.tsv file: $!\n";
open FAM, ">family_uterus.tsv" or die "Can't open the genus_uterus.tsv file: $!\n";
open ORDER, ">order_uterus.tsv" or die "Can't open the genus_uterus.tsv file: $!\n";

my ($count_g, $count_p, $count_c, $count_f, $count_o) = (1,1,1,1,1);

for my $n (0..35){
    my $ws_g = $wb_genus->add_worksheet(${$consolidated{'sample'}}[$n]);
    $ws_g->write(0,0,'Genus');
    $ws_g->write(0,1,'Number_16s');
    my $chart = $wb_genus->add_chart( type => 'column', embedded =>1 );
    print GENUS "Sample is ${$consolidated{'sample'}}[$n]\nGenus\tNumber_16s\n";
    my %genus_hash;
    for my $k (keys %consolidated){
	if ($k =~ /^genus/){
	    if(${$consolidated{$k}}[$n]>10){
		my @a = split("_",$k);
		$genus_hash{$a[1]}=${$consolidated{$k}}[$n];
	    }
	}
    }
    for my $v (sort {$genus_hash{$a} <=> $genus_hash{$b}} keys %genus_hash){	
	$ws_g->write($count_g,0,$v);
	$ws_g->write($count_g,1,$genus_hash{$v});
	$count_g++;
	print GENUS "$v\t$genus_hash{$v}\n";
    }
    $chart->add_series(
	name          => "Genus ${$consolidated{'sample'}}[$n]",
        categories    => '=\''.${$consolidated{'sample'}}[$n].'\'!$A$1:$A$'.$count_g,#xl_range_formula(${$consolidated{'sample'}}[$n], 1, $count_g, 0, 0 ),
        values        => '=\''.${$consolidated{'sample'}}[$n].'\'!$B$1:$B$'.$count_g #xl_range_formula(${$consolidated{'sample'}}[$n], 1, $count_g, 1, 1 )
    );
    $chart->set_legend (position => 'none');
    $chart->set_x_axis( name => 'Genus' );
    $chart->set_y_axis( name => 'Number 16s Sequences' );
    $ws_g->insert_chart( 'D2', $chart);
    $count_g=1;
    print GENUS "\n";    
#---------------------------------------------------------------------------------#

    my $ws_p = $wb_phylum->add_worksheet(${$consolidated{'sample'}}[$n]);
    $ws_p->write(0,0,'Phylum');
    $ws_p->write(0,1,'Number_16s');
    $chart = $wb_phylum->add_chart( type => 'column', embedded =>1 );
    print PHYL "Sample is ${$consolidated{'sample'}}[$n]\nPhylum\tNumber_16s\n";
    my %phylum_hash;
    print PHYL "Sample is ${$consolidated{'sample'}}[$n]\nPhylum\tNumber_16s\n";
    for my $k (keys %consolidated){
	if ($k =~ /^phylum/){
	    if(${$consolidated{$k}}[$n]>10){
		my @a = split("_",$k);
		$phylum_hash{$a[1]}=${$consolidated{$k}}[$n];
	    }
	}
    }
    for my $v (sort {$phylum_hash{$a} <=> $phylum_hash{$b}} keys %phylum_hash){	
	$ws_p->write($count_p,0,$v);
	$ws_p->write($count_p,1,$phylum_hash{$v});
	$count_p++;
	print PHYL "$v\t$phylum_hash{$v}\n";
    }
    $chart->add_series(
	name          => "Phylum ${$consolidated{'sample'}}[$n]",
        categories    => '=\''.${$consolidated{'sample'}}[$n].'\'!$A$1:$A$'.$count_p,#xl_range_formula(${$consolidated{'sample'}}[$n], 1, $count_g, 0, 0 ),
        values        => '=\''.${$consolidated{'sample'}}[$n].'\'!$B$1:$B$'.$count_p #xl_range_formula(${$consolidated{'sample'}}[$n], 1, $count_g, 1, 1 )
    );
    $chart->set_legend (position => 'none');
    $chart->set_x_axis( name => 'Phylum' );
    $chart->set_y_axis( name => 'Number 16s Sequences' );
    $ws_p->insert_chart( 'D2', $chart);
    $count_p=1;
    print PHYL "\n";
#---------------------------------------------------------------------------------#

    my $ws_c = $wb_class->add_worksheet(${$consolidated{'sample'}}[$n]);
    $ws_c->write(0,0,'Class');
    $ws_c->write(0,1,'Number_16s');
    $chart = $wb_class->add_chart( type => 'column', embedded =>1 );
    print CLASS "Sample is ${$consolidated{'sample'}}[$n]\nClass\tNumber_16s\n";
    my %class_hash;
    print CLASS "Sample is ${$consolidated{'sample'}}[$n]\nClass\tNumber_16s\n";
    for my $k (keys %consolidated){
	if ($k =~ /^class/){
	    if(${$consolidated{$k}}[$n]>10){
		my @a = split("_",$k);
		$class_hash{$a[1]}=${$consolidated{$k}}[$n];
	    }
	}
    }
    for my $v (sort {$class_hash{$a} <=> $class_hash{$b}} keys %class_hash){	
	$ws_c->write($count_c,0,$v);
	$ws_c->write($count_c,1,$class_hash{$v});
	$count_c++;
	print CLASS "$v\t$class_hash{$v}\n";
    }
    $chart->add_series(
	name          => "Class ${$consolidated{'sample'}}[$n]",
        categories    => '=\''.${$consolidated{'sample'}}[$n].'\'!$A$1:$A$'.$count_c,#xl_range_formula(${$consolidated{'sample'}}[$n], 1, $count_g, 0, 0 ),
        values        => '=\''.${$consolidated{'sample'}}[$n].'\'!$B$1:$B$'.$count_c #xl_range_formula(${$consolidated{'sample'}}[$n], 1, $count_g, 1, 1 )
    );
    $chart->set_legend (position => 'none');
    $chart->set_x_axis( name => 'Class' );
    $chart->set_y_axis( name => 'Number 16s Sequences' );
    $ws_c->insert_chart( 'D2', $chart);
    $count_c=1;
    print CLASS "\n";
#---------------------------------------------------------------------------------#

    my $ws_f = $wb_family->add_worksheet(${$consolidated{'sample'}}[$n]);
    $ws_f->write(0,0,'Family');
    $ws_f->write(0,1,'Number_16s');
    $chart = $wb_family->add_chart( type => 'column', embedded =>1 );
    print FAM "Sample is ${$consolidated{'sample'}}[$n]\nFamily\tNumber_16s\n";
    my %family_hash;
    #my $s = keys %consolidated;
    #print $s."\n";
    print FAM "Sample is ${$consolidated{'sample'}}[$n]\nFamily\tNumber_16s\n";
    for my $k (keys %consolidated){
	if ($k =~ /^family/){
	    if(${$consolidated{$k}}[$n]>10){
		my @a = split("_",$k);
		$family_hash{$a[1]}=${$consolidated{$k}}[$n];
	    }
	}
    }
    for my $v (sort {$family_hash{$a} <=> $family_hash{$b}} keys %family_hash){	
	$ws_f->write($count_f,0,$v);
	$ws_f->write($count_f,1,$family_hash{$v});
	$count_f++;
	print FAM "$v\t$family_hash{$v}\n";
    }
    $chart->add_series(
	name          => "Family ${$consolidated{'sample'}}[$n]",
        categories    => '=\''.${$consolidated{'sample'}}[$n].'\'!$A$1:$A$'.$count_f,#xl_range_formula(${$consolidated{'sample'}}[$n], 1, $count_g, 0, 0 ),
        values        => '=\''.${$consolidated{'sample'}}[$n].'\'!$B$1:$B$'.$count_f #xl_range_formula(${$consolidated{'sample'}}[$n], 1, $count_g, 1, 1 )
    );
    $chart->set_legend (position => 'none');
    $chart->set_x_axis( name => 'Family' );
    $chart->set_y_axis( name => 'Number 16s Sequences' );
    $ws_f->insert_chart( 'D2', $chart);
    $count_f=1;
    print FAM "\n";
#---------------------------------------------------------------------------------#
    my $ws_o = $wb_order->add_worksheet(${$consolidated{'sample'}}[$n]);
    $ws_o->write(0,0,'Order');
    $ws_o->write(0,1,'Number_16s');
    $chart = $wb_order->add_chart( type => 'column', embedded =>1 );
    print ORDER "Sample is ${$consolidated{'sample'}}[$n]\nOrder\tNumber_16s\n";
    my %order_hash;
    print ORDER "Sample is ${$consolidated{'sample'}}[$n]\nOrder\tNumber_16s\n";
    for my $k (keys %consolidated){
	if ($k =~ /^order/){
	    if(${$consolidated{$k}}[$n]>10){
		my @a = split("_",$k);
		$order_hash{$a[1]}=${$consolidated{$k}}[$n];
	    }
	}
    }
    for my $v (sort {$order_hash{$a} <=> $order_hash{$b}} keys %order_hash){	
	$ws_o->write($count_o,0,$v);
	$ws_o->write($count_o,1,$order_hash{$v});
	$count_o++;
	print ORDER "$v\t$order_hash{$v}\n";
    }
    $chart->add_series(
	name          => "Order ${$consolidated{'sample'}}[$n]",
        categories    => '=\''.${$consolidated{'sample'}}[$n].'\'!$A$1:$A$'.$count_o,#xl_range_formula(${$consolidated{'sample'}}[$n], 1, $count_g, 0, 0 ),
        values        => '=\''.${$consolidated{'sample'}}[$n].'\'!$B$1:$B$'.$count_o #xl_range_formula(${$consolidated{'sample'}}[$n], 1, $count_g, 1, 1 )
    );
    $chart->set_legend (position => 'none');
    $chart->set_x_axis( name => 'Order' );
    $chart->set_y_axis( name => 'Number 16s Sequences' );
    $ws_o->insert_chart( 'D2', $chart);
    $count_o=1;
    print ORDER "\n";
}

#print Dumper(%consolidated);
#print Dumper(%desig_data);
#print "Number of data messed up: $messed_data\n";
