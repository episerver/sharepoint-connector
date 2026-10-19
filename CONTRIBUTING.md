# Contributing

Thanks for your interest in improving **Optimizely Connect for SharePoint**!

This project is **community-maintained** and is **not** covered by Optimizely Support.
Optimizely retains control of the direction of the project and reserves the right to
close issues and pull requests that don't align with the project roadmap. The best way
to get a change in is to open an issue to discuss it first, then send a focused pull
request.

By contributing, you agree that your contributions are licensed under the project's
[Apache License 2.0](LICENSE) (see section 5, "Submission of Contributions"). No separate
contributor licence agreement (CLA) is required.

## Scope & compatibility

This addon targets **Optimizely CMS 11** on **.NET Framework 4.8**. There is no CMS 12/13
or SaaS CMS build. Please keep pull requests within that target — changes that assume
CMS 12+, .NET 5+, or SaaS CMS are out of scope.

## Reporting bugs and requesting features

- Search [existing issues](https://github.com/episerver/sharepoint-connector/issues) first
  to avoid duplicates.
- Open a new issue using the appropriate template and include steps to reproduce, the CMS
  version, and expected vs. actual behaviour.
- **Do not report security vulnerabilities in public issues.** Follow the process in
  [SECURITY.md](SECURITY.md) instead.

## Making changes

1. Fork the repository.
2. Create a feature branch from `master` (e.g. `bugfix/123-short-description` or
   `feature/123-short-description`, where `123` is the issue number).
3. Make your changes and add or update tests where appropriate.
4. Ensure the solution builds and all tests pass — see
   [Building](README.md#building) and [Running tests](README.md#running-tests).
5. Submit a pull request against `master`, referencing the related issue.

Please squash your changes into a single commit unless multiple commits make the history
clearer.

## Commit message guidelines

Based on [Chris Beams' seven rules](https://chris.beams.io/posts/git-commit/#seven-rules):

- Separate subject from body with a blank line
- Limit the subject line to ~50 characters
- Capitalize the subject line and don't end it with a period
- Use the imperative mood ("Add", "Fix", not "Added", "Fixes")
- Wrap the body at 72 characters
- Use the body to explain **what** and **why**, not **how**
- Reference the associated issue number

## Pull request expectations

- Keep the change focused; unrelated changes should be separate PRs.
- Make sure the code is documented and tests cover new behaviour.
- CI (build and tests) must be green before a PR can be merged.
- A maintainer from [@episerver/cms-addons](https://github.com/orgs/episerver/teams/cms-addons)
  will review; please be patient, as this is a community-maintained project.

## Code of conduct

This project follows the [Code of Conduct](CODE_OF_CONDUCT.md). By participating, you are
expected to uphold it.
