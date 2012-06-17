#import "WKMasterViewController.h"
#import "WKDetailViewController.h"
#import "WKDocReader.h"

@implementation WKMasterViewController

@synthesize detailViewController = _detailViewController;

- (id)initWithNibName:(NSString *)nibNameOrNil bundle:(NSBundle *)nibBundleOrNil
{
    self = [super initWithNibName:nibNameOrNil bundle:nibBundleOrNil];
    if (self) {
		self.title = NSLocalizedString(@"WKDocReader", @"WKDocReader");
    }
    return self;
}
							
- (void)dealloc
{
	[_detailViewController release];
    [super dealloc];
}

- (void)viewDidLoad
{
	NSArray *resources = [[NSBundle mainBundle] pathsForResourcesOfType:@"doc" inDirectory:@""];
	documents = [[NSMutableArray alloc] init];
	for(NSString *path in resources) {
		[documents addObject:[path lastPathComponent]];
	}
    [super viewDidLoad];
}

- (BOOL)shouldAutorotateToInterfaceOrientation:(UIInterfaceOrientation)interfaceOrientation
{
	return (interfaceOrientation != UIInterfaceOrientationPortraitUpsideDown);
}

- (NSInteger)numberOfSectionsInTableView:(UITableView *)tableView
{
	return 1;
}

- (NSInteger)tableView:(UITableView *)tableView numberOfRowsInSection:(NSInteger)section
{
	return [documents count];
}

- (UITableViewCell *)tableView:(UITableView *)tableView cellForRowAtIndexPath:(NSIndexPath *)indexPath
{
    static NSString *CellIdentifier = @"Cell";
    
    UITableViewCell *cell = [tableView dequeueReusableCellWithIdentifier:CellIdentifier];
    if (cell == nil) {
        cell = [[[UITableViewCell alloc] initWithStyle:UITableViewCellStyleDefault reuseIdentifier:CellIdentifier] autorelease];
        cell.accessoryType = UITableViewCellAccessoryDisclosureIndicator;
    }

	cell.textLabel.text = [documents objectAtIndex:indexPath.row];
    return cell;
}

- (void)tableView:(UITableView *)tableView didSelectRowAtIndexPath:(NSIndexPath *)indexPath
{
    if (!self.detailViewController) {
        self.detailViewController = [[[WKDetailViewController alloc] initWithNibName:@"WKDetailViewController" bundle:nil] autorelease];
    }
	NSString *filename = [documents objectAtIndex:indexPath.row];
	WKDocReader *reader = [[WKDocReader alloc] initWithContentsOfFile:[[NSBundle mainBundle] pathForResource:[filename stringByDeletingPathExtension] ofType:@"doc"]];
	
	self.detailViewController.detailItem = [reader attributedString];
	self.detailViewController.documentInfo = [reader documentAttributes];
	[reader release];
    [self.navigationController pushViewController:self.detailViewController animated:YES];
}

@end
