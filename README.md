# outlook-sig-comply-ad

Utility with a GUI interface that creates a user's Outlook signature by filling up a htm template with values obtained from an Active Directory serarch.

The utility allows the technician to adjust or add to the values obtained from the search and saves the user's signature to a network location.

Enforcing of the signature is done via an Active Directory group policy that runs the home/set-sig.vbs script at logon.

The set-sig.vbs script copies the signature to the local computer and sets it as the default signature in Outlook.
