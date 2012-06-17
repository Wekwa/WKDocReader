#import "WKDetailViewController.h"
#import "WKDocumentInfoViewController.h"

@interface WKDetailViewController ()
- (void)configureView;
@end

@implementation WKDetailViewController

@synthesize detailItem = _detailItem, documentInfo;

- (void)dealloc
{
	[_detailItem release];
    [super dealloc];
}

- (void)setDetailItem:(id)newDetailItem
{
        [_detailItem release]; 
        _detailItem = [newDetailItem retain]; 
		
        [self configureView];
    
}

- (void)configureView
{
	if (self.detailItem) {
		[coreTextView setAttributedString:self.detailItem];
	}
}

- (void)viewDidLoad
{
    [super viewDidLoad];
	coreTextView = [[WKCoreTextView alloc] initWithFrame:self.view.bounds];
	[self.view addSubview:coreTextView];
	self.navigationItem.rightBarButtonItem = infoButton;
	[self configureView];
}

- (BOOL)shouldAutorotateToInterfaceOrientation:(UIInterfaceOrientation)interfaceOrientation
{
	return (interfaceOrientation != UIInterfaceOrientationPortraitUpsideDown);
}

- (id)initWithNibName:(NSString *)nibNameOrNil bundle:(NSBundle *)nibBundleOrNil
{
    self = [super initWithNibName:nibNameOrNil bundle:nibBundleOrNil];
    if (self) {
		self.title = NSLocalizedString(@"Detail", @"Detail");
    }
    return self;
}

-(void)viewDocumentInfo:(id)sender
{
	WKDocumentInfoViewController *documentInfoController = [[WKDocumentInfoViewController alloc] init];
	[documentInfoController setDictionary:documentInfo];
	[self presentModalViewController:documentInfoController animated:YES];
	[documentInfoController release];
}
							
@end
