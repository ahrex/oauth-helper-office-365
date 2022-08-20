#!/usr/bin/env python3

# vim: set tw=88:
# match black's default line length for easier formatting:
# https://github.com/psf/black/blob/4ebf14d17ed544be893be5706c02116fd8b83b4c/src/black/const.py#L1

"""OAuth helper for Outlook 365"""

# inspired by the below; why not use that instead?
#
# https://www.vanormondt.net/~peter/blog/2021-03-16-mutt-office365-mfa.html#orgb39f967
#   * too rigid in implementation in the sense that it's specific to one institution
#
# https://github.com/UvA-FNWI/M365-IMAP
#   * good foundation, but requires hardcoding `config.py`
#   * could use this, and template out some hardcodes
#
# https://gitlab.com/muttmua/mutt/-/blob/master/contrib/mutt_oauth2.py.README
#   * OAuth flow hasn't been updated in a year+, and Outlook 365 changed their backend
#     such that it longer works
#   * previously tried w/o success to bring it up to par; also the code could use some
#     detangling
#
# given the above limitations, we'll roll our own implementation based off of the second
# source; by using Microsoft's own `msal` OAuth library, we should be future-proofing
# this a bit

import argparse
import getpass
import os
import re
import subprocess
import sys
import urllib.parse

try:
    import msal
except ModuleNotFoundError:
    sys.stderr.write(
        "missing required module `msal`, please install it from https://github.com/AzureAD/microsoft-authentication-library-for-python#installation\n"
    )
    sys.exit(1)


class MSOAuth(object):
    # https://docs.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth#get-an-access-token
    #
    # we want the `offline_access` scope as well, which is added by the `msal` library;
    # to confirm this, inspect the `scope` property of `initiate_auth_code_flow`'s
    # return
    scopes = [
        "https://outlook.office.com/IMAP.AccessAsUser.All",
        "https://outlook.office.com/SMTP.Send",
    ]

    # borrow Mozilla Thunderbird's client information
    # https://hg.mozilla.org/comm-central/file/e18ac7fbc3bdd18c46a12773732fa8c7c3360d9a/mailnews/base/src/OAuth2Providers.jsm#l129
    client_id = "08162f7c-0fd2-4200-a84a-f25a4db0b584"
    client_secret = "TxRBilcHdC6WGBee]fs?QR:SJ8nI[g82"

    def __init__(self, login_email, token_path, passphrase=None):
        self.token_path = token_path
        self.login_email = login_email
        self._passphrase = passphrase
        self.cache = msal.SerializableTokenCache()

    @staticmethod
    def ask_for_passphrase(prompt1, prompt2=None):
        while True:
            try:
                pp1 = getpass.getpass(prompt1)
                if not prompt2:
                    return pp1
                pp2 = getpass.getpass(prompt2)
            except KeyboardInterrupt:
                sys.stderr.write("\n")
                return
            if pp1 != pp2:
                sys.stderr.write("Passwords do not match, try again\n\n")
            else:
                return pp1

    @staticmethod
    def openssl_ops(passphrase, proc_args, stdin_bytes=None):
        """
        encrypt and decrypt using openssl
          * avoid gnupg, so we don't have to have a gpg keypair on disk
          * optionally send something via stdin (e.g. contents to be encrypted)
        """
        # create a pipe so we can pass the password somewhat securely
        # https://stackoverflow.com/a/48068951
        fd_read, fd_write = os.pipe()
        stdout = None
        stderr = None
        with subprocess.Popen(
            [
                "openssl",
                "enc",
                "-aes-256-cbc",
                "-pass",
                # https://xkcd.com/221/
                # yeah I have no idea how these numbers are determined; the below was
                # found through experimentation
                "fd:{!s}".format(4 if stdin_bytes else 3),
                "-e",
                "-salt",
                "-pbkdf2",
            ]
            + proc_args,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            pass_fds=(fd_read, fd_write),
        ) as p:
            os.write(fd_write, (passphrase + "\n").encode("utf-8"))
            try:
                stdout, stderr = p.communicate(stdin_bytes, timeout=10)
            except subprocess.TimeoutExpired:
                p.kill()
                stdout, stderr = p.communicate()
                sys.stderr.write("error: underlying encryption process timed out\n")
                return stdout, stderr, p.returncode
            finally:
                try:
                    os.close(fd_read)
                except Exception:
                    pass
                try:
                    os.close(fd_write)
                except Exception:
                    pass
            return stdout, stderr, p.returncode

    def _write_cache(self, passphrase):
        stdout, stderr, retcode = self.openssl_ops(
            passphrase,
            ["-out", self.token_path],
            self.cache.serialize().encode("utf-8"),
        )
        if retcode != 0 or len(stderr) > 0:
            sys.stderr.write(
                "error: did not successfully encrypt credentials to disk; stdout: {!r}, stderr: {!r}, return code {!r}\n".format(
                    stdout, stderr, retcode
                )
            )
            return False
        return True

    def save_new_token_to_disk(self):
        """
        request authorization from the OAuth endpoint, saving (+overwriting) the result
        onto disk
        """

        if not self._passphrase:
            passphrase = self.ask_for_passphrase(
                "Enter a passphrase we'll use to encrypt the token on disk: ",
                "Confirm the passphrase: ",
            )
            while passphrase == "":
                sys.stderr.write("\nwarning: passphrase cannot be blank, try again\n")
                passphrase = self.ask_for_passphrase(
                    "Enter a passphrase we'll use to encrypt the token on disk: ",
                    "Confirm the passphrase: ",
                )
            if passphrase is None:
                # possible `KeyboardInterrupt`, exit now
                return 1
            print()
        else:
            passphrase = self._passphrase

        # build our msal object
        #
        # reference:
        # https://msal-python.readthedocs.io/en/latest/#confidentialclientapplication
        #
        # inspiration:
        # https://github.com/AzureAD/microsoft-authentication-library-for-python/blob/dev/sample/confidential_client_secret_sample.py
        # https://github.com/UvA-FNWI/M365-IMAP/blob/main/get_token.py
        app = msal.ConfidentialClientApplication(
            self.client_id,
            client_credential=self.client_secret,
            token_cache=self.cache,
        )

        # reference:
        # https://msal-python.readthedocs.io/en/latest/#msal.ConfidentialClientApplication.initiate_auth_code_flow
        #
        # there's a `redirect_uri` allowlist for different `client_id` entries; if we
        # did not create our own `client_id`, we have to accomodate the `rediect_uri`
        # someone else has set
        result = app.initiate_auth_code_flow(self.scopes)

        # only check `scope` here because we're looking deeper to `result["scope"]`;
        # documentation says these fields are private, so we'll leave the rest of the
        # properties to be validated by `acquire_token_by_auth_code_flow` below
        expected_properties = ["scope"]
        missing_properties = []
        for expected_property in expected_properties:
            if expected_property not in result:
                missing_properties.append(expected_property)
        if len(missing_properties) > 0:
            sys.stderr.write(
                "error: missing properties `{!r}` in result: {!r}\n".format(
                    missing_properties, result
                )
            )
            return 1

        if "offline_access" not in result["scope"]:
            sys.stderr.write(
                "error: expected `offline_access` in `scope` but received this instead: {!r}\n".format(
                    result["scope"]
                )
            )
            return 1

        print("Please authenticate at: {}\n".format(result["auth_uri"]))
        try:
            response_uri = input("And paste the response URI here: ")
        except KeyboardInterrupt:
            sys.stderr.write("\n")
            return 1
        print()

        parsed_qs = None

        try:
            parsed_uri = urllib.parse.urlparse(response_uri)
            parsed_qs = urllib.parse.parse_qs(parsed_uri.query)
        except Exception as e:
            sys.stderr.write("error: could not parse URI: {!s}\n".format(e))
            return 1

        if not parsed_qs:
            sys.stderr.write(
                "error: did not successfully parse the query string from URI {!r}\n".format(
                    response_uri
                )
            )
            return 1

        if "code" not in parsed_qs:
            sys.stderr.write(
                "error: did not successfully find a code from query string {!r}\n".format(
                    parsed_qs
                )
            )
            return 1

        if len(parsed_qs["code"]) != 1:
            sys.stderr.write(
                "error: got an unexpected number of codes in {!r}\n".format(
                    parsed_qs["code"]
                )
            )
            return 1

        # we now have a code in `parsed_qs["code"][0]`; finish the auth flow by
        # redeeming it for a long-lived `access_token`
        #
        # fix the following library error by tweaking our input data:
        #
        #     ValueError: state mismatch: NdrIBvgcWylRzPQk vs ['NdrIBvgcWylRzPQk']
        #
        # any further state mismatches could be a user error, so let's still catch 'em
        try:
            result = app.acquire_token_by_auth_code_flow(
                {**result, **{"state": [result["state"]]}},
                parsed_qs,
                scopes=self.scopes,
            )
        except ValueError as e:
            m = re.search(r"^state mismatch: \['([^']+)'\] vs \['([^']+)'\]$", str(e))
            if m:
                internal_state, user_provided_state = m.groups()
                sys.stderr.write(
                    "error: internal state {!r} does not match user-provided state {!r}; perhaps you pasted an old response?\n".format(
                        internal_state, user_provided_state
                    )
                )
                return 1
            # not our error, re-raise
            raise

        if "access_token" not in result or "error" in result:
            sys.stderr.write(
                "error: did not successfully fetch access_token: {!s}\n".format(
                    result.get("error", "(unknown)")
                )
            )
            sys.stderr.write("\n{!s}\n".format(result.get("error_description", "")))
            return 1

        # if we're here without error, our access token should be in
        # `result["access_token"]`, as well as in the cache; let's encrypt and write the
        # cache to disk, per
        # https://msal-python.readthedocs.io/en/latest/#msal.SerializableTokenCache
        if not self._write_cache(passphrase):
            return 1

        print("Successfully saved access token to {!r}".format(self.token_path))
        return 0

    def print_token_from_disk(self):
        """
        decrypt a previous authorization from disk, possibly refreshing it + saving the
        refreshed copy, and returing the token on stdout for consumption
        """

        if not os.path.exists(self.token_path):
            sys.stderr.write(
                "error: path {!r} does not exist\n".format(self.token_path)
            )
            return 1

        if not self._passphrase:
            passphrase = self.ask_for_passphrase("Enter the decryption passphrase: ")
        else:
            passphrase = self._passphrase

        if passphrase is None:
            # possible `KeyboardInterrupt`, exit now
            return 1
        elif passphrase == "":
            sys.stderr.write("error: aborting on empty passphrase\n")
            return 1

        stdout, stderr, retcode = self.openssl_ops(
            passphrase, ["-d", "-in", self.token_path]
        )

        if retcode != 0 or len(stderr) > 0:
            sys.stderr.write(
                "error: failed to decrypt token file {!r}; stderr: {!r}, return code {!r}\n".format(
                    self.token_path, stderr, retcode
                )
            )
            return 1

        # reconstitute our cache and application
        # https://msal-python.readthedocs.io/en/latest/#msal.SerializableTokenCache
        self.cache.deserialize(stdout.decode("utf-8"))
        app = msal.ConfidentialClientApplication(
            self.client_id,
            client_credential=self.client_secret,
            token_cache=self.cache,
        )

        accounts = app.get_accounts(username=self.login_email)
        if len(accounts) != 1:
            # TODO: handle multiple accounts for a username
            sys.stderr.write(
                "error: got an unexpected number of accounts {!r}\n".format(accounts)
            )
            return 1

        cache_before = app.token_cache.serialize()
        # this method possibly refreshes the token via
        # `_acquire_token_silent_from_cache_and_possibly_refresh_it`; we should write
        # back a refreshed version to disk, if it was updated
        #
        # https://github.com/AzureAD/microsoft-authentication-library-for-python/blob/a18c2231896d8a050ad181461928f4dbd818049f/sample/confidential_client_secret_sample.py#L56-L59
        result = app.acquire_token_silent(self.scopes, accounts[0])
        cache_after = app.token_cache.serialize()

        if not result:
            sys.stderr.write(
                "error: no suitable token exists in cache, please re-authorize via `--mode authorize`\n"
            )
            return 1

        if cache_before != cache_after:
            if not self._write_cache(passphrase):
                sys.stderr.write(
                    "error: weird how we could read the cache, but not write it\n"
                )
                return 1

        if "access_token" not in result or "error" in result:
            sys.stderr.write(
                "error: did not successfully fetch access_token: {!s}\n".format(
                    result.get("error", "(unknown)")
                )
            )
            sys.stderr.write("\n{!s}\n".format(result.get("error_description", "")))
            return 1

        print(result["access_token"])
        return 0


def main(argv):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--mode",
        default="mission",
        help='whether we should "authorize" an access token, or use/refresh a token in "mission" mode',
    )
    parser.add_argument(
        "--passphrase",
        help="passphrase to encrypt/decrypt, useful for non-interactive sessions (e.g. within mutt)",
    )
    parser.add_argument("username", help="username to authenticate as")
    parser.add_argument("token_path", help="path to store the encrypted token")

    args = parser.parse_args()
    oauth_handler = MSOAuth(args.username, args.token_path, passphrase=args.passphrase)

    if args.mode == "mission":
        return oauth_handler.print_token_from_disk()
    elif args.mode == "authorize":
        return oauth_handler.save_new_token_to_disk()
    else:
        sys.stderr.write("error: unknown mode {!r}\n".format(args.mode))
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
