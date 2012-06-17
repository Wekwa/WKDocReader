#import "WKDocumentInfoViewController.h"

@implementation WKDocumentInfoViewController

@synthesize dictionary;

-(void)viewDidLoad
{
	for(id key in [dictionary allKeys]) {
		NSString *string = [textView text];
		string = [string stringByAppendingFormat:@"%@: %@\n", key, [[dictionary objectForKey:key] description]];
		[textView setText:string];
	}
}

-(IBAction)donePressed:(id)sender
{
	[self dismissModalViewControllerAnimated:YES];
}

@end
