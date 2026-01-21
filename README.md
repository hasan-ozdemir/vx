# VX (VS2022)

Local CLI that attaches to running Visual Studio 2022 instances via the ROT (no extension required).

The earlier VSIX prototype remains in `src/VxRemoteControl` for in-proc automation experiments, but the primary entrypoint is the CLI in `src/vx`.

## Build

```
dotnet build VxRemoteControl.sln
```

## Usage

```
dotnet run --project src/vx -- info
dotnet run --project src/vx -- open solution C:\path\to\MySolution.sln
```

Options:
- `--all` show all VS2022 instances
- `--instance N` pick a specific instance index
- `--json` machine-readable output

## Info output

`vx info` includes:
- VS running?
- Version and edition
- Solution name/path + active configuration
- Project list with kinds
- Open documents + active document

`vx open solution <path>` opens the solution in the active VS2022 instance (or starts a new instance if none are running).

