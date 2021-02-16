package Excel::Grinder;

our $VERSION = "1.0";

# time to grow up
use strict;
use warnings;
use Carp;

# this stands on the feet of giants
use Excel::Writer::XLSX;
use Spreadsheet::XLSX;

# OO out of habit
sub new {
	my ($class, $default_directory) = @_;
	
	# default the default directory to /tmp/excel_grinder
	$default_directory ||= '/tmp/excel_grinder';
	
	# make sure that directory exists
	mkdir $default_directory if !(-d $default_directory);
	
	# if it still does exist, bail out
	carp "Error: $default_directory does not exist and cannot be auto-created." if !(-d $default_directory);
	
	# become!
	my $self = bless {
		'default_directory' => $default_directory,
	}, $class;
	
	return $self;
}

# method to convert a three-level array into a nice excel file
sub write_excel {
	# required arguments are (1) the filename and (2) the data structure to turn into an XLSX file
	my ($self, %args) = @_;
	# looks like:
	#	'filename' => 'some_file.xlsx', # will be saved under /opt/majestica/tmp/DATABASE_NAME/some_file.xlsx; required
	#	'the_data' => @$three_level_arrayref, # worksheets->rows->columns; see below; required
	#	'headings_in_data' => 1, # if filled, first row of each worksheet will be captialized; optional
	#	'worksheet_names' => ['Names','of','Worksheets'], # if filled, will be the names to give the worksheets

	my ($tmp_dir, $item, $col, @bits, $workbook, $worksheet_data, $worksheet, $n, $row_array, $row_upper, $worksheet_name);

	# fail without a filename
	carp 'Error: Filename required for write_excel()' if !$args{filename};

	# the data structure must be an array of arrays of arrays
	# three levels: worksheets, rows, columns
	carp 'Error: Must send a three-level arrayref (workbook->rows->columns) to write_excel()' if !$args{the_data}[0][0][0];

	# place into default_directory unless they specified a directory

	$args{filename} = $self->default_directory.'/'.$args{filename} if $args{filename} !~ /\//;
	$args{filename} .= '.xlsx' if $args{filename} !~ /\.xlsx$/;

	# start our workbook
	$workbook = Excel::Writer::XLSX->new( $args{filename} );

	# Set the format for dates.
	my $date_format = $workbook->add_format( num_format => 'mm/dd/yy' );

	# start adding worksheets
	foreach $worksheet_data (@{ $args{the_data} }) {
		$worksheet_name = shift @{ $args{worksheet_names} }; # if it's there
		$worksheet_name =~ s/[^0-9a-z\-\s]//gi; # clean it up
		$worksheet = $workbook->add_worksheet($worksheet_name);
		
		# go thru each row...
		$n = 0;
		foreach $row_array (@$worksheet_data) {

			# do they want the first row to the headings?
			if ($args{headings_in_data} && $n == 0) { # uppercase the first row
				@$row_upper = map { uc($_) } @$row_array;
				$row_array = $row_upper;
			}
			
			# now each column...
			$col = 0;
			foreach $item (@$row_array) {
				# dates are no funzies
				if ($item =~ /^(\d{4})-(\d{2})-(\d{2})$/) { # special routine for dates
					$worksheet->write_date_time( $n, $col++, $1.'-'.$2.'-'.$3.'T', $date_format );
				} else {
					 $worksheet->write( $n, $col++, $item );
				}

			}
			$n++;
		}
	}

	# that's not so hard, now is it?
	return $args{filename};
}

# method to import an excel file into a nice three-level array
sub read_excel {
	# require argument is the filename or  full path to the excel xlsx file
	# if it's just a filename, look in the default directory
	my ($self,$filename) = @_;

	$filename = $self->{default_directory}.'/'.$filename if $filename !~ /\//;	
	$filename .= '.xlsx' if $filename !~ /\.xlsx$/;

	# gotta exist, after all that
	carp 'Error: Must send a valid full file path to an XLSX file to read_excel()' if !(-e "$filename");
	
	my ($excel, $sheet_num, $sheet, $row_num, $row, @the_data, $cell, $col);

	# again, stand on the shoulders of giants
	$excel = Spreadsheet::XLSX->new($filename);

	# read it in, sheet by sheet
	$sheet_num = 0;
	foreach $sheet (@{$excel->{Worksheet}}) {

		# set the max = 0 if there is one or none rows
		$sheet->{MaxRow} ||= $sheet->{MinRow};

		# same for the columns
		$sheet->{MaxCol} ||= $sheet->{MinCol};

		# cycle through each row
		$row_num = 0;
		foreach $row ($sheet->{MinRow} .. $sheet->{MaxRow}) {
			# go through each available column
			foreach $col ($sheet->{MinCol} ..  $sheet->{MaxCol}) {

                # get ahold of the actual cell object
				$cell = $sheet->{Cells}[$row][$col];
				
				# next if !$cell; # skip if blank

				# add it to our nice array
				push (@{ $the_data[$sheet_num][$row] }, $cell->{Val} );
			}
			# advance
			$row_num++;
        }
		$sheet_num++;
	}

	# send it back
	return \@the_data;
}

1;

__END__

=head1 Excel::Grinder

This allows Majestica to export data to Excel files, as well as to convert Excel files into 
data structures.  That second function is to allow for some nasty batch-update tools.  
Please note that for both of these, we are only supporting the 'modern' XLSX format.

Start it up like so, passing in a valid UtilityBelt object:

	$xlsx = Excel::Grinder->new('/default/directory/for/your/excel/files');

=head2 write_excel()

To write out an XLSX file, you will prepare a nice three-level arrayref.
The actual data is at the third level; the first two are organizational 
to represent worksheets and rows.

Here is a nice example:

	$full_file_path = $xlsx->write_excel(
		'filename' => 'our_family.xlsx',
		'headings_in_data' => 1,
		'worksheet_names' => ['Dogs','People'],
		'the_data' => [
			[
				['Name','Main Trait','Age Type'],
				['Ginger','Wonderful','Old'],
				['Pepper','Loving','Passed'],
				['Polly','Fun','Young']
				['Daisy','Crazy','Puppy']
			],
			[
				['Name','Main Trait','Age Type'],
				['Melanie','Smart','Oldish'],
				['Lorelei','Fun','Young'],
				['Eric','Fat','Old']
			]
		],
	);

That will create a file at /opt/majestica/tmp/DATABASE_NAME/ginger.xlsx .  Please use
your file_manager object if you need to save it to a proper spot.  The
$full_file_path variable is now your full path to the file on the disk.

In normal use, you'd probably prepare the arrayref beforehand and then
call like so:

	$xlsx->write_excel(
		'filename' => 'ginger.xlsx',
		'the_data' => \@my_data,
		'headings_in_data' => 1,
		'worksheet_names' => ['Dogs','People'],
	);

The 'headings_in_data' arg tells use to make each worksheet's first row
all caps to indicate those are the headings.  Fancy.  We are not exactly
stretching the use of Excel::Writer::XLSX here.

The 'worksheet_names' argument is the arrayref to the names to put on the
nice tabs for the worksheets.  Both 'worksheet_names' and 'headings_in_data'
are optional.


=head2 read_excel()

This does the exact opposite of write_excel() in that it reads in an XLSX
file and returns the arrayref in the exact same format as what write_excel()
receives.  All it needs is the absolute filepath for an XLSX file:

	$the_data = $xlsx->read_excel('/opt/majestica/tmp/DATABASE_NAME/ginger.xlsx');

@$the_data will look like the structure in the example above.  Try it out ;)
