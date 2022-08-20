oauth-helper-office-365
====

Plug-and-play OAuth helper for Office 365

Operates in two modes:
  * authorize mode: perform the first half of the OAuth dance, prompting the
    user to log in, and encrypting the returned secret to disk using openssl
  * mission mode: decrypt the secret from disk, printing to stdout a JWT that
    can be used to validate the previous authorization

# Requirements

* [microsoft authentication library
  (msal)](https://github.com/AzureAD/microsoft-authentication-library-for-python)
  - `pip install msal` if you receive a "missing required module `msal`" error
* [openssl](https://www.openssl.org/source/)

# Usage

## Barebones usage

The below example is a smoke test of functionality. In order for the script to
be useful, you'll want to consume the mission mode token somewhere.

### Generate and save secret

1. Run the authorization command
```
my_user_login@computer:~/oauth-helper-office-365$ python3 oauth-helper-office-365.py --mode authorize my_user_login@contoso.com ./secret-token.bin
```

1. Enter a passphrase on stdin (optional: pass the passphrase as an argument
   above)
```
Enter a passphrase we'll use to encrypt the token on disk:
Confirm the passphrase:
```

1. Open the authentication URL when prompted
```
Please authenticate at: https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=<snip>&response_type=code&scope=https%3A%2F%2Foutlook.office.com%2FIMAP.AccessAsUser.All+https%3A%2F%2Foutlook.office.com%2FSMTP.Send+offline_access+openid+profile&state=<snip>&code_challenge=<snip>&code_challenge_method=S256&nonce=<snip>&client_info=1
```

1. Once authenticated, your browser will try to open a URL at `localhost`. Paste
   that URL back into the prompt
```
And paste the response URI here: http://localhost/?code=<snip>&client_info=<snip>&state=<snip>&session_state=<snip>#
```

1. Validate that the token has been saved
```
Successfully saved access token to './secret-token.bin'
```

### Print a JWT from the saved secret

1. Run the mission mode command, entering your previous passphrase when prompted
```
my_user_login@computer:~/oauth-helper-office-365$ python3 oauth-helper-office-365.py my_user_login@contoso.com ./secret-token.bin
Enter the decryption passphrase:
```

1. See the JWT that's generated
```
eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiaWF0IjoxNTE2MjM5MDIyfQ.SflKxwRJSMeKKF2QT4fwpMeJf36POk6yJV_adQssw5c
```

## Mutt usage

[examples/muttrc](examples/muttrc) should contain enough mutt configuration
decrypt a stored token. The below guide will help you encrypt and store a secret
to allow the `muttrc` to work.

### Directories and includes

1. Generate the directory structure where we'll store our secret
```
mkdir -p ~/.config/mutt
```

1. Include the [example `muttrc`](examples/muttrc) in your config, changing
   `imap_user` as necessary

1. Optional: uncomment the last bits of the example muttrc to standardize folder
   names and mail deletion policies in mutt

### Generate and save secret

1. Run the authorization command
```
my_user_login@computer:~/oauth-helper-office-365$ python3 oauth-helper-office-365.py --mode authorize my_user_login@contoso.com ~/.config/mutt/office-365.token
```

1. Enter a passphrase on stdin (optional: pass the passphrase as an argument
   above)
```
Enter a passphrase we'll use to encrypt the token on disk:
Confirm the passphrase:
```

1. Open the authentication URL when prompted
```
Please authenticate at: https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=<snip>&response_type=code&scope=https%3A%2F%2Foutlook.office.com%2FIMAP.AccessAsUser.All+https%3A%2F%2Foutlook.office.com%2FSMTP.Send+offline_access+openid+profile&state=<snip>&code_challenge=<snip>&code_challenge_method=S256&nonce=<snip>&client_info=1
```

1. Once authenticated, your browser will try to open a URL at `localhost`. Paste
   that URL back into the prompt
```
And paste the response URI here: http://localhost/?code=<snip>&client_info=<snip>&state=<snip>&session_state=<snip>#
```

1. Validate that the token has been saved
```
Successfully saved access token to './secret-token.bin'
```

### Run mutt with the new config

1. Run mutt as usual, entering your previous passphrase when prompted
```
my_user_login@computer:~$ mutt
```
