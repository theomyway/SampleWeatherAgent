# Overview of the Employee Onboarding Agent (POC)

This project is a proof-of-concept Employee Onboarding Agent built with the [Microsoft 365 Agents SDK](https://github.com/Microsoft/Agents).

The agent guides users through a simple onboarding happy flow using:
- Text prompts for guidance
- Adaptive Cards for structured question and answer collection

The current implementation uses in-memory conversation state (no database), which is suitable for local demos and POC scenarios.

## Onboarding happy flow

The agent collects the following information in sequence:
1. Full name
2. Work email
3. Department
4. Equipment needs (multi-select)
5. Policy acknowledgement (Yes/No)
6. Final onboarding summary card

After completion, users can type `restart` to begin the flow again.

## Get started with the template

> **Prerequisites**
>
> To run locally, you need:
>
> - [Node.js](https://nodejs.org/) (supported versions: 18, 20, 22)
> - [Microsoft 365 Agents Toolkit VS Code extension](https://aka.ms/teams-toolkit) latest version or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

> For local debugging using the Microsoft 365 Agents Toolkit CLI, follow:
> [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. Open the Microsoft 365 Agents Toolkit panel in VS Code.
2. Press F5 and select `Debug in Microsoft 365 Agents Playground`.
3. Send any message (for example, `hi`) to start onboarding.

## What's included in this project

| Folder | Contents |
| - | - |
| `.vscode` | VS Code debug files |
| `appPackage` | App manifest and packaging assets |
| `env` | Environment files |
| `infra` | Templates for provisioning Azure resources |
| `src` | Application source code |

Key files:

| File | Contents |
| - | - |
| `src/index.ts` | Starts the agent server |
| `src/agent.ts` | Main onboarding conversation logic (happy flow state machine) |
| `src/tools/*.ts` | Optional helper tools (not required by current onboarding flow) |

Project configuration files:

| File | Contents |
| - | - |
| `m365agents.yml` | Main Microsoft 365 Agents Toolkit project file |
| `m365agents.local.yml` | Overrides for local execution and debugging |
| `m365agents.playground.yml` | Overrides for Microsoft 365 Agents Playground debugging |

## Notes and limitations (POC)

- Session state is in memory and resets when the process restarts.
- No persistent storage or external workflow system is configured.
- Validation is intentionally simple and demo-focused.

## Next suggested enhancements

- Stronger input validation and formatting rules
- HR system integrations (ticketing, approvals, provisioning endpoints)
- Durable state persistence for production
- Branching onboarding paths (contractor vs full-time, department-specific journeys)

## Additional references

- [Microsoft 365 Agents Toolkit documentation](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit samples](https://github.com/OfficeDev/TeamsFx-Samples)
