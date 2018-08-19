use strict;

use OpenOffice::OODoc;

my ($doc, $body, @list, $ltxt, $ltxt);

my $sd = "odt.out";
my $of;

opendir(my $dh, $sd) || die;

foreach my $d (grep { /[^\.]/ }
               readdir $dh) {

    print "$d\n";
    
    opendir($dh, "$sd/$d/") || die;
     
    foreach (grep { /\.odt$/ }
             readdir $dh) {

        print "\t$_\n";
        
        do_conv($d, $_);
    }


}

sub do_conv {
    my ($d, $fn) = @_;

    open($of, ">", "$sd/$d/$fn.txt") or die;
    
    $doc = odfDocument(file => "$sd/$d/$fn");

    $doc->outputDelimitersOff();

    $body = $doc->getBody();
    @list = $body->selectChildElements(".*");

    $ltxt = undef;

    foreach (@list) {
        #print $_->tag();
        #print "\n";
        my $tn = $_->tag();
     
        if ($tn =~ /text:p/) {
            # может быть пустым или содержать
            # или произвольное количество записей
            do_p($_);
        } elsif ($tn =~ /text:soft-page-break/) {
            # новая колонка
            #print "lb\n";
            print $of "<br>\n";
        } else {
            print "**** new tag: $tn ****\n";
            exit(1);
        }
        
    }
}

sub do_p {
    my @list = $_[0]->selectChildElements(".*");
    
    return unless @list;
    
    #print "do_p\n";
    
    print $of "<br>\n" unless $ltxt =~ '^$';
    
    foreach (@list) {
        my $tn  = $_->tag();
        my $txt = $doc->getText($_);
        
        $ltxt = $txt;
        
        if ($txt =~ '^$') { # новая запись
            print $of "<br>\n";
        } else {
            my $st = $doc->textStyle($_);
            my %sp = $doc->styleProperties($st);

            my $fs = $sp{'fo:font-style'};
            my $fw = $sp{'fo:font-weight'};

            #print "\t\t $tn $txt /$fs, $fw/\n"; next;
            
            if ($fs eq 'italic') {
                print $of "<i>";
            } 
            if ($fw eq 'bold') {
                print $of "<b>";
            }
            
            print $of "$txt";
            
            if ($fs eq 'italic') {
                print $of "<\/i>";
            } 
            if ($fw eq 'bold') {
                print $of "<\/b>";
            } 
        }
    }
}