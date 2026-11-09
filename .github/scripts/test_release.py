import contextlib
from io import StringIO
import json
import os
from pathlib import Path
import plistlib
import shutil
import subprocess
import sys
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
METADATA_FIXTURES = {
    release.PROJECT: "\n".join(
        f"config{index} = {{ buildSettings = {{ "
        "MARKETING_VERSION = 1.0.7; CURRENT_PROJECT_VERSION = 1.0.7; }; };"
        for index in range(6)
    ).encode(),
    release.PLIST: plistlib.dumps({"CFBundleShortVersionString": "1.0.7"}),
    release.LOCALIZED: (
        'CFBundleName = "msoffice crypt";\n'
        'CFBundleShortVersionString = "1.0.7";\n'
        'CFBundleGetInfoString = "msoffice crypt version 1.0.7, Copyright 2021 miyako.";\n'
    ).encode(),
}


def not_found():
    return HTTPError("https://api.github.com", 404, "Not Found", {}, None)


def api_response(data, status=200):
    response = StringIO(json.dumps(data))
    response.status = status
    return response


CREATED_RELEASE = {"id": 123, "tag_name": "v1.0.8", "draft": False, "prerelease": False}


class MetadataTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.previous = Path.cwd()
        os.chdir(self.temp.name)
        self.addCleanup(os.chdir, self.previous)
        self.reset_metadata()

    def reset_metadata(self, *paths):
        for path in paths or METADATA_FIXTURES:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(METADATA_FIXTURES[path])

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

    def test_fixture_version_and_no_added_build_key(self):
        self.assertEqual(release.read_version(), "1.0.7")
        self.assertNotIn("CFBundleVersion", plistlib.loads(release.PLIST.read_bytes()))

    def test_all_bump_modes_coordinate_every_field(self):
        for mode, expected in (("patch", "1.0.8"), ("minor", "1.1.0"), ("major", "2.0.0")):
            with self.subTest(mode=mode):
                self.reset_metadata()
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
            self.reset_metadata()

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
                    self.reset_metadata(release.PLIST)
                    self.set_plist(**{key: value})
                    with self.assertRaises(ValueError):
                        release.read_version()

    def test_xcode_drift_and_missing_settings_rejected(self):
        for key in release.SETTINGS:
            for replacement in (f"{key} = 1.0.4;", "", f'{key} = "$(UNKNOWN)";'):
                with self.subTest(key=key, replacement=replacement):
                    self.reset_metadata(release.PROJECT)
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
            self.reset_metadata(release.LOCALIZED)
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


class LiveRepositoryTests(unittest.TestCase):
    def setUp(self):
        previous = Path.cwd()
        os.chdir(ROOT)
        self.addCleanup(os.chdir, previous)

    def test_repository_metadata_consistent(self):
        marketing = release.assignments(release.PROJECT.read_text(), "MARKETING_VERSION")
        self.assertTrue(marketing)
        expected = release.require_version(release.setting_value(marketing[0]))
        self.assertEqual(release.read_version(), expected)

    def test_regression_suite_passes_after_future_bump_in_copy(self):
        before = {path: path.read_bytes() for path in METADATA_FIXTURES}
        current = release.read_version()
        with tempfile.TemporaryDirectory() as directory:
            copied = Path(directory)
            shutil.copytree(ROOT / ".github", copied / ".github",
                            ignore=shutil.ignore_patterns("__pycache__"))
            for path, content in before.items():
                target = copied / path
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(content)
            environment = dict(os.environ, GITHUB_OUTPUT="", PYTHONDONTWRITEBYTECODE="1",
                               PYTHONPATH=str(copied / ".github/scripts"))
            bumped = subprocess.run(
                [sys.executable, ".github/scripts/release.py", "bump", "major"],
                cwd=copied, env=environment, text=True, capture_output=True, check=True,
            ).stdout.strip()
            self.assertNotEqual(bumped, current)
            release.require_version(bumped)
            # Run every regression group and the live consistency check in the copy,
            # excluding only this recursive copy-proof test.
            result = subprocess.run(
                [sys.executable, "-m", "unittest", "test_release.MetadataTests",
                 "test_release.PublicationTests", "test_release.WorkflowTests",
                 "test_release.LiveRepositoryTests.test_repository_metadata_consistent"],
                cwd=copied, env=environment, text=True, capture_output=True, check=False,
            )
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(before, {path: path.read_bytes() for path in before})


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
    def test_authenticated_empty_listing_allows_publication(self, run, urlopen):
        run.return_value.returncode = 2
        urlopen.side_effect = [
            api_response({"full_name": "example/plugin"}), api_response([]),
        ]
        release.ensure_available("v1.0.7")
        self.assertEqual([entry.args[0].full_url for entry in urlopen.call_args_list], [
            "https://api.github.com/repos/example/plugin",
            "https://api.github.com/repos/example/plugin/releases?per_page=100&page=1",
        ])
        for entry in urlopen.call_args_list:
            request = entry.args[0]
            self.assertEqual(request.get_method(), "GET")
            self.assertEqual(request.get_header("Authorization"), "Bearer test-token")

    @patch("release.urlopen")
    @patch("release.subprocess.run")
    def test_published_and_draft_collisions_beyond_first_page(self, run, urlopen):
        run.return_value.returncode = 2
        for draft in (False, True):
            with self.subTest(draft=draft):
                urlopen.side_effect = [
                    api_response({"full_name": "example/plugin"}),
                    api_response([{"tag_name": f"v0.0.{i}", "draft": False} for i in range(100)]),
                    api_response([{"tag_name": "v1.0.7", "draft": draft}]),
                ]
                with self.assertRaisesRegex(ValueError, "Release .* already exists"):
                    release.ensure_available("v1.0.7")
                self.assertTrue(urlopen.call_args.args[0].full_url.endswith("page=2"))

    @patch("release.urlopen")
    @patch("release.subprocess.run")
    def test_all_release_pages_checked_before_absence(self, run, urlopen):
        run.return_value.returncode = 2
        urlopen.side_effect = [
            api_response({"full_name": "example/plugin"}),
            api_response([{"tag_name": f"v0.0.{i}", "draft": True} for i in range(100)]),
            api_response([]),
        ]
        release.ensure_available("v1.0.7")
        self.assertEqual(urlopen.call_count, 3)

    @patch("release.urlopen")
    @patch("release.subprocess.run")
    def test_api_errors_never_mean_absence(self, run, urlopen):
        run.return_value.returncode = 2
        for stage in ("repository", "listing", "later-page"):
            for error in (not_found(), HTTPError("url", 401, "Unauthorized", {}, None),
                          HTTPError("url", 403, "Forbidden", {}, None),
                          HTTPError("url", 500, "Server error", {}, None), URLError("unreachable")):
                responses = []
                if stage != "repository":
                    responses.append(api_response({"full_name": "example/plugin"}))
                if stage == "later-page":
                    responses.append(api_response([{"tag_name": "v0.0.1", "draft": True}] * 100))
                urlopen.side_effect = [*responses, error]
                with self.subTest(stage=stage, error=error), self.assertRaises((HTTPError, URLError)):
                    release.ensure_available("v1.0.7")

    @patch("release.urlopen")
    @patch("release.subprocess.run")
    def test_malformed_repository_and_listings_fail_closed(self, run, urlopen):
        run.return_value.returncode = 2
        for repository in (None, [], {}, {"full_name": "another/repo"}):
            urlopen.side_effect = [api_response(repository)]
            with self.subTest(repository=repository), self.assertRaises(ValueError):
                release.ensure_available("v1.0.7")
        for listing in (None, {}, [None], [{}], [{"tag_name": 7, "draft": True}],
                        [{"tag_name": "v1.0.7"}], [{"tag_name": "", "draft": False}],
                        [{"tag_name": "v1.0.7", "draft": "false"}]):
            urlopen.side_effect = [
                api_response({"full_name": "example/plugin"}), api_response(listing),
            ]
            with self.subTest(listing=listing), self.assertRaises(ValueError):
                release.ensure_available("v1.0.7")
        invalid_json = StringIO("not JSON")
        invalid_json.status = 200
        urlopen.side_effect = [api_response({"full_name": "example/plugin"}), invalid_json]
        with self.assertRaises(ValueError):
            release.ensure_available("v1.0.7")
        urlopen.side_effect = [api_response({"full_name": "example/plugin"}, status=202)]
        with self.assertRaisesRegex(RuntimeError, "Unexpected GitHub API status"):
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
        self.api = stack.enter_context(patch("release.urlopen"))
        self.api.return_value = api_response(CREATED_RELEASE, status=201)
        return available, output, run

    def test_publish_rechecks_then_tags_exact_sha_and_never_clobbers(self):
        available, output, run = self.publication_context()
        release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        available.assert_called_once_with("v1.0.8")
        self.assertEqual(run.call_args_list, [
            call(["git", "tag", "v1.0.8", SHA], check=True),
            call(["git", "-c", "push.followTags=false", "push", "origin",
                  "refs/tags/v1.0.8:refs/tags/v1.0.8"], check=True),
            call(["gh", "release", "upload", "v1.0.8", "plugin.zip", "plugin.dmg",
                  "--repo", "example/plugin"], check=True),
        ])
        self.api.assert_called_once()
        request = self.api.call_args.args[0]
        self.assertEqual(request.full_url, "https://api.github.com/repos/example/plugin/releases")
        self.assertEqual(request.get_method(), "POST")
        self.assertEqual(json.loads(request.data), {
            "tag_name": "v1.0.8", "target_commitish": SHA, "name": "v1.0.8",
            "draft": False, "prerelease": False, "generate_release_notes": True,
        })
        self.assertEqual(request.get_header("Authorization"), "Bearer test-token")
        self.assertEqual(self.api.call_args.kwargs, {"timeout": 30})

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
        self.api.assert_not_called()

    def test_remote_tag_mismatch_never_creates_release(self):
        _, output, run = self.publication_context()
        output.return_value = f"{'b' * 40}\trefs/tags/v1.0.8\n"
        with self.assertRaisesRegex(ValueError, "does not point"):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        self.assertEqual(run.call_count, 2)
        self.api.assert_not_called()

    def test_release_creation_failure_is_not_swallowed_or_retried(self):
        _, _, run = self.publication_context()
        for error in (not_found(), HTTPError("url", 422, "Unprocessable", {}, None),
                      HTTPError("url", 500, "Server error", {}, None), URLError("timed out")):
            with self.subTest(error=error):
                run.reset_mock()
                self.api.reset_mock()
                self.api.side_effect = error
                with self.assertRaisesRegex(RuntimeError, "retained; no rollback or retry"):
                    release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
                self.assertEqual(run.call_count, 2)
                self.api.assert_called_once()

    def test_uncertain_create_response_never_uploads_or_rolls_back(self):
        _, _, run = self.publication_context()
        for response in (None, {}, dict(CREATED_RELEASE, id=True),
                         dict(CREATED_RELEASE, tag_name="v9.0.0"),
                         dict(CREATED_RELEASE, draft=True), dict(CREATED_RELEASE, prerelease=True)):
            with self.subTest(response=response):
                run.reset_mock()
                self.api.reset_mock()
                self.api.return_value = api_response(response, status=201)
                with self.assertRaisesRegex(RuntimeError, "Inspect .* manually"):
                    release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
                self.assertEqual(run.call_count, 2)
                self.api.assert_called_once()
        self.api.return_value = api_response(CREATED_RELEASE, status=202)
        with self.assertRaisesRegex(RuntimeError, "Unexpected GitHub API status"):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        invalid_json = StringIO("response lost")
        invalid_json.status = 201
        self.api.return_value = invalid_json
        with self.assertRaisesRegex(RuntimeError, "result is uncertain"):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])

    def test_upload_error_is_not_retried_or_rolled_back(self):
        _, _, run = self.publication_context()
        run.side_effect = [None, None, subprocess.CalledProcessError(
            1, ["gh", "release", "upload"], stderr="connection lost",
        )]
        with self.assertRaisesRegex(RuntimeError, "Inspect .* manually"):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        self.assertEqual(run.call_count, 3)
        self.api.assert_called_once()
        self.assertEqual(run.call_args.args[0], [
            "gh", "release", "upload", "v1.0.8", "plugin.zip", "plugin.dmg",
            "--repo", "example/plugin",
        ])

    def test_partial_upload_failure_keeps_created_release_and_assets(self):
        _, _, run = self.publication_context()
        retained = {"tag": False, "release": False, "assets": []}

        def create(request, **kwargs):
            self.assertEqual(request.get_method(), "POST")
            self.assertEqual(run.call_count, 2)
            retained["release"] = True
            return api_response(CREATED_RELEASE, status=201)

        def command(args, **kwargs):
            if args[:2] == ["git", "-c"]:
                retained["tag"] = True
            if args[:3] == ["gh", "release", "upload"]:
                retained["assets"].append("plugin.zip")
                raise subprocess.CalledProcessError(1, args, stderr="second upload failed")

        self.api.side_effect = create
        run.side_effect = command
        with self.assertRaisesRegex(RuntimeError, "retained; no rollback or retry"):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        self.assertEqual(retained, {"tag": True, "release": True, "assets": ["plugin.zip"]})
        self.assertEqual(run.call_count, 3)
        self.api.assert_called_once()
        self.assertNotIn("--clobber", run.call_args.args[0])

    def test_ambiguous_post_keeps_server_created_release_without_retry(self):
        _, _, run = self.publication_context()
        retained = []

        def create(request, **kwargs):
            retained.append(json.loads(request.data))
            raise URLError("connection lost after server created release")

        self.api.side_effect = create
        with self.assertRaisesRegex(RuntimeError, "result is uncertain"):
            release.publish("1.0.8", ["plugin.zip", "plugin.dmg"])
        self.assertEqual(len(retained), 1)
        self.assertEqual(run.call_count, 2)
        self.api.assert_called_once()

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
                with patch("release.urlopen", side_effect=[
                    api_response({"full_name": "example/plugin"}), api_response([]),
                ]):
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
