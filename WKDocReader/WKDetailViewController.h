#import <UIKit/UIKit.h>
#import "WKCoreTextView.h"

@interface WKDetailViewController : UIViewController {
	WKCoreTextView *coreTextView;
	IBOutlet UIBarButtonItem *infoButton;
}

@property (strong, nonatomic) id detailItem;
@property (strong, nonatomic) NSDictionary *documentInfo;

-(IBAction)viewDocumentInfo:(id)sender;

@end
