import { ActivityTypes } from "@microsoft/agents-activity";
import {
  AgentApplicationBuilder,
  MessageFactory,
  TurnContext,
} from "@microsoft/agents-hosting";

export const onboardingAgent = new AgentApplicationBuilder().build();

onboardingAgent.onConversationUpdate(
  "membersAdded",
  async (context: TurnContext) => {
    await context.sendActivity(
      "Hello and welcome! I can guide you through a quick employee onboarding process."
    );
  }
);

type OnboardingStep =
  | "start"
  | "fullName"
  | "email"
  | "department"
  | "equipment"
  | "policy"
  | "complete";

interface OnboardingSession {
  step: OnboardingStep;
  fullName?: string;
  email?: string;
  department?: string;
  equipment?: string[];
  policyAccepted?: boolean;
}

const onboardingSessions = new Map<string, OnboardingSession>();

function getSession(conversationId: string): OnboardingSession {
  const existing = onboardingSessions.get(conversationId);
  if (existing) {
    return existing;
  }

  const created: OnboardingSession = {
    step: "start",
  };

  onboardingSessions.set(conversationId, created);
  return created;
}

function createQuestionCard(
  title: string,
  question: string,
  inputId: string,
  options?: string[],
  isMultiSelect?: boolean
) {
  const body: any[] = [
    {
      type: "TextBlock",
      text: title,
      weight: "Bolder",
      size: "Medium",
      wrap: true,
    },
    {
      type: "TextBlock",
      text: question,
      wrap: true,
      spacing: "Small",
    },
  ];

  if (options && options.length > 0) {
    body.push({
      type: "Input.ChoiceSet",
      id: inputId,
      style: "expanded",
      isMultiSelect: Boolean(isMultiSelect),
      choices: options.map((opt) => ({ title: opt, value: opt })),
    });
  } else {
    body.push({
      type: "Input.Text",
      id: inputId,
      placeholder: "Type your answer",
    });
  }

  return {
    type: "AdaptiveCard",
    version: "1.5",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    body,
    actions: [
      {
        type: "Action.Submit",
        title: "Submit",
      },
    ],
  };
}

function createSummaryCard(session: OnboardingSession) {
  return {
    type: "AdaptiveCard",
    version: "1.5",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    body: [
      {
        type: "TextBlock",
        text: "Onboarding Complete ✅",
        weight: "Bolder",
        size: "Medium",
      },
      {
        type: "FactSet",
        facts: [
          { title: "Name", value: session.fullName ?? "-" },
          { title: "Email", value: session.email ?? "-" },
          { title: "Department", value: session.department ?? "-" },
          {
            title: "Equipment",
            value: session.equipment?.length ? session.equipment.join(", ") : "-",
          },
          {
            title: "Policy Accepted",
            value: session.policyAccepted ? "Yes" : "No",
          },
        ],
      },
    ],
  };
}

async function sendCard(context: TurnContext, cardContent: any) {
  const response = MessageFactory.attachment({
    contentType: "application/vnd.microsoft.card.adaptive",
    content: cardContent,
  });
  await context.sendActivity(response);
}

onboardingAgent.onActivity(ActivityTypes.Message, async (context) => {
  try {
    const conversationId = context.activity.conversation?.id;
    if (!conversationId) {
      await context.sendActivity("Unable to identify conversation. Please try again.");
      return;
    }

    const session = getSession(conversationId);
    const submittedData = (context.activity.value ?? {}) as Record<string, string>;
    const text = (context.activity.text ?? "").trim();

    if (session.step === "start") {
      await context.sendActivity("Great! Let’s begin onboarding. First question:");
      await sendCard(
        context,
        createQuestionCard(
          "Employee Onboarding",
          "What is your full name?",
          "fullName"
        )
      );
      session.step = "fullName";
      return;
    }

    if (session.step === "fullName") {
      const fullName = submittedData.fullName ?? text;
      if (!fullName) {
        await context.sendActivity("Please provide your full name.");
        return;
      }

      session.fullName = fullName;
      session.step = "email";
      await context.sendActivity("Thanks. Next:");
      await sendCard(
        context,
        createQuestionCard("Contact", "What is your work email?", "email")
      );
      return;
    }

    if (session.step === "email") {
      const email = submittedData.email ?? text;
      if (!email || !email.includes("@")) {
        await context.sendActivity("Please provide a valid email address.");
        return;
      }

      session.email = email;
      session.step = "department";
      await context.sendActivity("Perfect. Choose your department:");
      await sendCard(
        context,
        createQuestionCard(
          "Department",
          "Select your department.",
          "department",
          ["Engineering", "HR", "Finance", "Sales", "Operations"]
        )
      );
      return;
    }

    if (session.step === "department") {
      const department = submittedData.department ?? text;
      if (!department) {
        await context.sendActivity("Please choose a department.");
        return;
      }

      session.department = department;
      session.step = "equipment";
      await context.sendActivity("Now pick required equipment (you can choose multiple):");
      await sendCard(
        context,
        createQuestionCard(
          "Equipment",
          "Which equipment do you need?",
          "equipment",
          ["Laptop", "Monitor", "Keyboard", "Mouse", "Headset"],
          true
        )
      );
      return;
    }

    if (session.step === "equipment") {
      const rawEquipment = submittedData.equipment ?? text;
      if (!rawEquipment) {
        await context.sendActivity("Please select at least one equipment item.");
        return;
      }

      session.equipment = rawEquipment.split(",").map((v) => v.trim()).filter(Boolean);
      session.step = "policy";
      await context.sendActivity("Final step: policy confirmation.");
      await sendCard(
        context,
        createQuestionCard(
          "Policy Acknowledgement",
          "Do you accept the company onboarding policies?",
          "policyAccepted",
          ["Yes", "No"]
        )
      );
      return;
    }

    if (session.step === "policy") {
      const policyAnswer = (submittedData.policyAccepted ?? text).toLowerCase();
      if (!policyAnswer) {
        await context.sendActivity("Please answer policy acknowledgement.");
        return;
      }

      session.policyAccepted = policyAnswer === "yes";
      session.step = "complete";

      await context.sendActivity("Thanks! Here is your onboarding summary:");
      await sendCard(context, createSummaryCard(session));
      await context.sendActivity("If you want to restart onboarding, type: restart");
      return;
    }

    if (session.step === "complete") {
      if (text.toLowerCase() === "restart") {
        onboardingSessions.set(conversationId, { step: "start" });
        await context.sendActivity("Onboarding restarted.");
      } else {
        await context.sendActivity(
          "Onboarding is already complete. Type 'restart' to begin again."
        );
      }
    }
  } catch (err: any) {
    try {
      console.error(
        "Error handling onboarding activity:",
        JSON.stringify(err, Object.getOwnPropertyNames(err), 2)
      );
    } catch (e) {
      console.error("Error handling onboarding activity (fallback):", err);
    }

    if (err instanceof Error) {
      throw err;
    }

    throw new Error(typeof err === "string" ? err : JSON.stringify(err));
  }
});
