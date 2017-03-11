#Reveal URL Domain in New Mail in Outlook 2007

### Background

When reading an email, either in its own window or in the reading pane, sometimes Outlook has difficulties displaying
the actual URL destination tooltip.  This is a MAJOR SECURITY concern in preventing PHISHING attacks.

Often, after finding the perfect location to hover the mouse over the URL, the tooltip simply flashes up and then disappears
before it can be read.  At other times it works ok.  Some emails are more problematic than others, and this might have something
to do with their content/mark up.

### The Code

As new emails arrive in your Inbox, the code intercepts them and searches for anchors in the email.  The destination URL is
copied into the visible URL text. This may make some pages look poor, but that's a small price to pay for peace of mind.  
At the moment, only anchors that use `://` to specify a URL are handled (with and without a URI component), and malformed URLs are ignored.
Other anchor types are ignored, including `mailto:`.  
The code is executed after any rules are processed, thus emails destined to other mail folders (in accordance to existing rules)
are not processed.  

It is possible, and planned, to make the code run against emails in the *current* folder, etc. to allow processing existing mail.

### Outlook Versions

The code in the repository is primarily aimed at Outlook 2007, though it may work and may be useful for other Outlook versions.
The code, separated in the repository into two files, is written in VBA and both methods should be placed into the one
`ThisOutlookSession` visual basic file that is available via `Tools -> Macro -> Visual Basic Editor`.

#### Note
*I am far from an expert on developing scripting for Outlook let alone Office, so any feedback on better methods, techniques,
and practices is most welcome.*




