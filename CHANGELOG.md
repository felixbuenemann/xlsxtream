# Changelog

## 1.3.2 (2017-10-30)

- Fix circular require in workbook
- Fix wrong escaping of extended Unicode characters

## 1.3.1 (2017-10-30)

- Restore stringio require in workbook

## 1.3.0 (2017-10-30)

- Drop rubyzip buffering workarounds, require rubyzip >= 1.2.0 (#17)
- Drop Ruby 1.9.1 compatibility (rubyzip 1.2 requires ruby >= 1.9.2)
- Refactor IO wrappers (#18)

## 1.2.0 (2017-10-30)

- Add support for customizing default font (#17)

## 1.1.0 (2017-10-23)

- Add support for boolean values (#13)

## 1.0.1 (2017-10-22)

- Fix writing unnamed worksheets with options

## 1.0.0 (2017-10-22)

- Don't close IO objects passed to Workbook constructor (#12)

## 0.3.1 (2017-10-22)

- Escape invalid XML 1.0 characters (#11)

## 0.3.0 (2017-07-12)

- Add support for auto-formatting

## 0.2.0 (2017-02-20)

- Ruby 2.4 compatibility
- Misc bug fixes

## 0.1.0 (2015-10-17)

- Initial release
