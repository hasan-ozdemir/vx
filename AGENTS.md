# Repository Guidelines

## Project Structure & Module Organization
- `src/vx/` contains the CLI application (entry point: `Program.cs`). This is the primary tool users run.
- `src/VxRemoteControl/` contains the earlier VSIX prototype used for in‑process automation experiments.
- `README.md` documents usage and examples.
- Build artifacts are ignored (`bin/`, `obj/`, `.vs/`).

## Build, Test, and Development Commands
- `dotnet build VxRemoteControl.sln` — builds the full solution (CLI + VSIX).
- `dotnet build src/vx/vx.csproj -c Release` — builds only the CLI.
- `dotnet run --project src/vx -- info` — runs the CLI and prints VS2022 status.
- `dotnet run --project src/vx -- open solution C:\path\to\MySolution.sln` — opens a solution in VS2022.

## Coding Style & Naming Conventions
- Language: C#; target framework `net8.0-windows` for the CLI.
- Indentation: 4 spaces, braces on new lines (default C# style).
- Naming: `PascalCase` for types/methods, `camelCase` for locals/fields, constants in `PascalCase`.
- Keep changes minimal and favor small, composable helper methods (e.g., `TryGet`, `EnumerateProjects`).

## Testing Guidelines
- No automated tests are currently defined in this repository.
- For changes that affect behavior, verify manually using `vx info` and `vx open solution` against a running VS2022 instance.

## Commit & Pull Request Guidelines
- Commit messages follow an imperative, concise style (e.g., “Add vx open solution command”).
- Keep commits focused: one logical change per commit.
- PRs should include a short summary, reproduction steps (if applicable), and any relevant screenshots for UI changes.

## Agent-Specific Instructions
- If any file changes are made during a task, ensure changes are committed and pushed before responding with a summary.
