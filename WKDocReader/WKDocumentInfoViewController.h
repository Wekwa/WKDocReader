#import <UIKit/UIKit.h>

@interface WKDocumentInfoViewController : UIViewController {
	IBOutlet UITextView *textView;
}

-(IBAction)donePressed:(id)sender;

@property (nonatomic, strong) NSDictionary *dictionary;

@end
