# Corpus Validation

Tests use the Open XML SDK asset corpus at commit
`cd2b359ef824737edb93f1c6157c19551aae1e52`. Expected failures are listed in
`Benchmarks/upstream-corpus-expected-errors.tsv`.

Place the pinned SDK checkout next to OfficeKit, then run:

```sh
swift build -c release
.build/release/OfficeKitBenchmarks \
  --maximum-seconds 30 \
  --all-xml \
  --expected-errors Benchmarks/upstream-corpus-expected-errors.tsv \
  ../Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets
```

Mutation tests check bounded handling of invalid input.
