import contextlib
from io import StringIO
import os
from pathlib import Path
import plistlib
import shutil
import subprocess
import tempfile
import unittest
from unittest.mock import call, patch
from urllib.error import HTTPError, URLError

import release


ROOT = Path(__file__).resolve().parents[2]
SHA = "a" * 40
API_ENV = {
    "GITHUB_API_URL": "https://api.github.com",
    "GITHUB_REPOSITORY": "example/plugin",
    "GH_TOKEN": "test-token",
    "GITHUB_SHA": SHA,
}


def not_found():
    return HTTPError("https://api.github.com", 404, "Not Found", {}, None)


class MetadataTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.previous = Path.cwd()
        os.chdir(self.temp.name)
        self.addCleanup(os.chdir, self.previous)
        for path in (release.PROJECT, release.PLIST, release.LOCALIZED):
            path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(ROOT / path, path)

    def replace(self, path, old, new):
        text = path.read_text()
        self.assertIn(old, text)
        path.write_text(text.replace(old, new))

    def set_plist(self, **values):
        with release.PLIST.open("rb") as source:
            data = plistlib.load(source)
        data.update(values)
        with release.PLIST.open("wb") as target:
            plistlib.dump(data, target)

    def test_repository_version_and_no_added_build_key(self):
        self.assertEqual(release.read_version(), "1.0.7")
        self.assertNotIn("CFBundleVersion", plistlib.loads(release.PLIST.read_bytes()))

    def test_all_bump_modes_coordinate_every_field(self):
        for mode, expected in (("patch", "1.0.8"), ("minor", "1.1.0"), ("major", "2.0.0")):
            with self.subTest(mode=mode):
                for path in (release.PROJECT, release.PLIST, release.LOCALIZED):
                    shutil.copyfile(ROOT / path, path)
                self.assertEqual(release.bump(mode), expected)
                self.assertEqual(release.read_version(), expected)
                for key in release.SETTINGS:
                    self.assertEqual(release.assignments(release.PROJECT.read_text(), key), [expected] * 6)
                self.assertNotIn("CFBundleVersion", plistlib.loads(release.PLIST.read_bytes()))

    def test_known_variables_preserved_with_optional_build_literal(self):
        for reference in ("$(MARKETING_VERSION)", "${MARKETING_VERSION}"):
            self.set_plist(CFBundleShortVersionString=reference, CFBundleVersion="1.0.7")
            self.assertEqual(release.bump("patch"), "1.0.8")
            data = plistlib.loads(release.PLIST.read_bytes())
            self.assertEqual(data["CFBundleShortVersionString"], reference)
            self.assertEqual(data["CFBundleVersion"], "1.0.8")
            for path in (release.PROJECT, release.PLIST, release.LOCALIZED):
                shutil.copyfile(ROOT / path, path)

    def test_both_build_references_preserved(self):
        self.set_plist(
            CFBundleShortVersionString="$(MARKETING_VERSION)",
            CFBundleVersion="${CURRENT_PROJECT_VERSION}",
        )
        self.assertEqual(release.bump("minor"), "1.1.0")
        self.assertEqual(plistlib.loads(release.PLIST.read_bytes())["CFBundleVersion"],
                         "${CURRENT_PROJECT_VERSION}")

    def test_plist_drift_and_unknown_variables_rejected(self):
        for key in release.PLIST_SETTINGS:
            for value in ("1.0.4", "$(UNKNOWN)", "$(CURRENT_PROJECT_VERSION)" if
                          key == "CFBundleShortVersionString" else "$(MARKETING_VERSION)", "", 107):
                with self.subTest(key=key, value=value):
                    shutil.copyfile(ROOT / release.PLIST, release.PLIST)
                    self.set_plist(**{key: value})
                    with self.assertRaises(ValueError):
                        release.read_version()

    def test_xcode_drift_and_missing_settings_rejected(self):
        for key in release.SETTINGS:
            for replacement in (f"{key} = 1.0.4;", "", f'{key} = "$(UNKNOWN)";'):
                with self.subTest(key=key, replacement=replacement):
                    shutil.copyfile(ROOT / release.PROJECT, release.PROJECT)
                    text = release.PROJECT.read_text()
                    release.PROJECT.write_text(text.replace(f"{key} = 1.0.7;", replacement, 1))
                    with self.assertRaises(ValueError):
                        release.read_version()

    def test_invalid_semver_rejected(self):
        for value in ("01.0.7", "1.00.7", "1.0.07", "v1.0.7", "1.0", "1.0.7-beta", "1.0.7+1", "", "1.0.7\n"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                release.require_version(value)

    def test_invalid_xcode_literal_not_skipped(self):
        self.replace(release.PROJECT, "MARKETING_VERSION = 1.0.7;", "MARKETING_VERSION = 01.0.7;")
        with self.assertRaisesRegex(ValueError, "Invalid semantic"):
            release.read_version()

    def test_quoted_xcode_values_and_invalid_quoting(self):
        for key in release.SETTINGS:
            self.replace(release.PROJECT, f"{key} = 1.0.7;", f'{key} = "1.0.7";')
        self.assertEqual(release.read_version(), "1.0.7")
        self.replace(release.PROJECT, 'MARKETING_VERSION = "1.0.7";',
                     'MARKETING_VERSION = "1.0.7;')
        with self.assertRaisesRegex(ValueError, "Invalid semantic"):
            release.read_version()

    def test_localized_drift_rejected(self):
        for old in ('"1.0.7"', "version 1.0.7,"):
            shutil.copyfile(ROOT / release.LOCALIZED, release.LOCALIZED)
            self.replace(release.LOCALIZED, old, old.replace("1.0.7", "1.0.1"))
            with self.assertRaisesRegex(ValueError, "Localized"):
                release.read_version()

    def test_drift_blocks_bump_without_writing(self):
        self.replace(release.PROJECT, "CURRENT_PROJECT_VERSION = 1.0.7;",
                     "CURRENT_PROJECT_VERSION = 1.0.4;")
        before = {path: path.read_bytes() for path in
                  (release.PROJECT, release.PLIST, release.LOCALIZED)}
        with self.assertRaises(ValueError):
            release.bump("patch")
        self.assertEqual(before, {path: path.read_bytes() for path in before})

    def test_built_plist_rejects_unresolved_variables_and_drift(self):
        release.validate_plist({"CFBundleShortVersionString": "1.0.7"}, "1.0.7", False)
        for value in ("$(MARKETING_VERSION)", "1.0.4"):
            with self.assertRaises(ValueError):
                release.validate_plist({"CFBundleShortVersionString": value}, "1.0.7", False)

    def test_output_format(self):
        output = Path("outputs")
        with patch.dict(os.environ, {"GITHUB_OUTPUT": str(output)}), \
                patch("sys.argv", ["release.py", "version"]), contextlib.redirect_stdout(StringIO()):
            release.main()
        self.assertEqual(output.read_text(), "version=1.0.7\ntag=v1.0.7\n")


class PublicationTests(unittest.TestCase):
    def setUp(self):
        self.environment = patch.dict(os.environ, API_ENV)
        self.environment.start()
        self.addCleanup(self.environment.stop)

    @patch("release.urlopen")
    @patch("release.subprocess.run")
    def test_existing_remote_tag_rejected_before_api(self, run, urlopen):
        run.return_value.returncode = 0
        with self.assertRaisesRegex(ValueError, "already exists"):
            release.ensure_available("v1.0.7")
        urlopen.assert_not_called()
        self.assertEqual(run.call_args.args[0][-1], "refs/tags/v1.0.7")

    @patch("release.urlopen")
    @patch("release.subprocess.run")
    def test_remote_failure_is_not_absence(self, run, urlopen):
        run.return_value.returncode = 128
        with self.assertRaisesRegex(RuntimeError, "Cannot check"):
            release.ensure_available("v1.0.7")
        urlopen.assert_not_called()

    @patch("release.urlopen")
    @patch("release.subprocess.run")
    def test_only_api_404_allows_publication(self, run, urlopen):
        run.return_value.returncode = 2
        urlopen.side_effect = not_found()
        release.ensure_available("v1.0.7")
        self.assertTrue(urlopen.call_args.args[0].full_url.endswith("/releases/tags/v1.0.7"))
        urlopen.side_effect = None
        with self.assertRaisesRegex(ValueError, "Release .* already exists"):
            release.ensure_available("v1.0.7")
        for error in (HTTPError("url", 403, "Forbidden", {}, None),
                      HTTPError("url", 500, "Server error", {}, None), URLError("unreachable")):
            urlopen.side_effect = error
            with self.subTest(error=error), self.assertRaises((HTTPError, URLError)):
                release.ensure_available("v1.0.7")

    @patch("release.subprocess.check_output")
    def test_checkout_must_match_selected_sha(self, output):
        output.return_value = SHA + "\n"
        self.assertEqual(release.check_commit(), SHA)
        output.return_value = "b" * 40
        with self.assertRaisesRegex(ValueError, "does not match"):
            release.check_commit()
        with patch.dict(os.environ, {"GITHUB_SHA": "main"}), self.assertRaises(ValueError):
            release.check_commit()

    def publication_context(self):
        stack = contextlib.ExitStack()
        self.addCleanup(stack.close)
        stack.enter_context(patch("release.check_commit", return_value=SHA))
        stack.enter_context(patch("release.read_version", return_value="1.0.8"))
        available = stack.enter_context(patch("release.ensure_available"))
        stack.enter_context(patch("release.Path.is_file", return_value=True))
        stat = stack.enter_context(patch("release.Path.stat"))
        stat.return_value.st_size = 100
        output = stack.enter_context(patch("release.subprocess.check_output"))
        output.return_value = f"{SHA}\trefs/tags/v1.0.8\n"
        run = stack.enter_context(patch("release.subprocess.run"))
        return available, output, run

    def test_publish_rechecks_then_tags_exact_sha_and_never_clobbers(self):
        available, output, run = self.publication_context()
        release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        available.assert_called_once_with("v1.0.8")
        self.assertEqual(run.call_args_list, [
            call(["git", "tag", "v1.0.8", SHA], check=True),
            call(["git", "-c", "push.followTags=false", "push", "origin",
                  "refs/tags/v1.0.8:refs/tags/v1.0.8"], check=True),
            call(["gh", "release", "create", "v1.0.8", "--verify-tag", "--target", SHA,
                  "--title", "v1.0.8", "--generate-notes", "plugin.zip", "plugin.dmg"], check=True),
        ])

    def test_publication_rejection_never_tags(self):
        available, _, run = self.publication_context()
        available.side_effect = ValueError("existing tag or release")
        with self.assertRaises(ValueError):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        run.assert_not_called()

    def test_missing_asset_or_version_drift_never_tags(self):
        _, _, run = self.publication_context()
        with self.assertRaises(ValueError):
            release.publish("1.0.9", ["plugin.zip", "plugin.dmg"])
        with patch("release.Path.is_file", return_value=False), self.assertRaises(ValueError):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        run.assert_not_called()

    def test_failed_tag_push_never_creates_release(self):
        _, _, run = self.publication_context()
        run.side_effect = [None, subprocess.CalledProcessError(1, ["git", "push"])]
        with self.assertRaises(subprocess.CalledProcessError):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        self.assertEqual(run.call_count, 2)

    def test_remote_tag_mismatch_never_creates_release(self):
        _, output, run = self.publication_context()
        output.return_value = f"{'b' * 40}\trefs/tags/v1.0.8\n"
        with self.assertRaisesRegex(ValueError, "does not point"):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        self.assertEqual(run.call_count, 2)

    def test_release_creation_failure_is_not_swallowed_or_retried(self):
        _, _, run = self.publication_context()
        run.side_effect = [None, None, subprocess.CalledProcessError(1, ["gh", "release", "create"])]
        with self.assertRaises(subprocess.CalledProcessError):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        self.assertEqual(run.call_count, 3)

    @patch("release.ensure_available")
    @patch("release.check_commit", return_value=SHA)
    @patch("release.read_version", return_value="1.0.8")
    def test_preflight_validates_commit_and_availability(self, version, commit, available):
        with patch("sys.argv", ["release.py", "preflight"]), \
                patch.dict(os.environ, {"GITHUB_OUTPUT": ""}), contextlib.redirect_stdout(StringIO()):
            release.main()
        version.assert_called_once_with()
        commit.assert_called_once_with()
        available.assert_called_once_with("v1.0.8")

    def test_real_git_remote_tag_collision(self):
        with tempfile.TemporaryDirectory() as directory:
            remote = Path(directory) / "remote.git"
            checkout = Path(directory) / "checkout"
            subprocess.run(["git", "init", "--bare", "--quiet", str(remote)], check=True)
            subprocess.run(["git", "init", "--quiet", str(checkout)], check=True)
            previous = Path.cwd()
            try:
                os.chdir(checkout)
                subprocess.run(["git", "remote", "add", "origin", str(remote)], check=True)
                subprocess.run(["git", "-c", "user.name=Test", "-c", "user.email=test@example.com",
                                "commit", "--allow-empty", "--quiet", "-m", "fixture"], check=True)
                with patch("release.urlopen", side_effect=not_found()):
                    release.ensure_available("v1.0.8")
                subprocess.run(["git", "tag", "v1.0.8"], check=True)
                subprocess.run(["git", "push", "--quiet", "origin", "refs/tags/v1.0.8"], check=True)
                with patch("release.urlopen") as api, self.assertRaisesRegex(ValueError, "already exists"):
                    release.ensure_available("v1.0.8")
                api.assert_not_called()
            finally:
                os.chdir(previous)


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.build = (ROOT / ".github/workflows/release.yml").read_text()
        self.bump = (ROOT / ".github/workflows/bump-version.yml").read_text()

    def test_manual_build_no_inputs_or_tag_trigger(self):
        self.assertIn("name: Build\n", self.build)
        trigger = self.build.split("\non:\n", 1)[1].split("\npermissions:", 1)[0]
        self.assertEqual(trigger.strip(), "workflow_dispatch:")
        self.assertNotIn("inputs.", self.build)

    def test_all_build_checkouts_are_immutable(self):
        self.assertEqual(self.build.count("uses: actions/checkout@v4"), 3)
        self.assertEqual(self.build.count("ref: ${{ github.sha }}"), 3)
        self.assertNotIn("ref: ${{ needs.version.outputs.tag }}", self.build)

    def test_publication_waits_for_required_checks(self):
        self.assertIn("needs: [version, build-windows]", self.build)
        self.assertIn("group: release-publication\n  cancel-in-progress: false", self.build)
        publish = self.build.index("release.py publish")
        for gate in ("release.py preflight", "verify windows dependencies", "verify universal binary",
                     "release.py verify-bundle", "codesign --verify",
                     'status == "Accepted"', 'xcrun stapler validate "$DMG_PATH"'):
            self.assertLess(self.build.index(gate), publish)
        self.assertNotIn("always()", self.build)
        self.assertNotIn("continue-on-error", self.build)
        self.assertNotIn("softprops/action-gh-release", self.build)

    def test_bump_commits_only_selected_branch(self):
        self.assertIn("ref: ${{ github.sha }}", self.bump)
        self.assertIn('"HEAD:refs/heads/${GITHUB_REF_NAME}"', self.bump)
        self.assertIn("push.followTags=false", self.bump)
        self.assertNotIn("git tag", self.bump)
        self.assertNotIn("refs/tags/", self.bump)
        self.assertIn('"$GITHUB_REF" != refs/heads/*', self.bump)
        for mode in ("patch", "minor", "major"):
            self.assertIn(f"- {mode}", self.bump)


if __name__ == "__main__":
    unittest.main()
