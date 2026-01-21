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
dotnet run --project src/vx -- view MainPage.xaml.cs
dotnet run --project src/vx -- view !MyProject:MainPage.xaml.cs
dotnet run --project src/vx -- ls us
dotnet run --project src/vx -- ls cl
dotnet run --project src/vx -- ls ns
dotnet run --project src/vx -- ls mt:MainPage
dotnet run --project src/vx -- build
dotnet run --project src/vx -- rebuild
dotnet run --project src/vx -- clean
dotnet run --project src/vx -- build bot*
dotnet run --project src/vx -- build !*ios
dotnet run --project src/vx -- deploy !*ios
dotnet run --project src/vx -- !BotyumApp:build
dotnet run --project src/vx -- !BotyumApp:rebuild
dotnet run --project src/vx -- !BotyumApp:clean
dotnet run --project src/vx -- !BotyumApp:deploy
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

`vx view <filename>` opens the first matching file in the active solution.
`vx view !ProjectName:<filename>` opens the file within a specific project.

`vx ls us|cl|ns` lists using directives, class names, or namespaces from the active document.
`vx ls mt:ClassName` lists methods under the specified class in the active document.

`vx build`, `vx rebuild`, and `vx clean` operate on the active solution. Add a project pattern to target the first match (wildcard `*` supported).
`vx deploy !projectPattern` deploys the first matching project (wildcard `*` supported).
`vx !ProjectName:build|rebuild|clean|deploy` runs the action for a specific project in the active solution (wildcard `*` supported in the project name).

