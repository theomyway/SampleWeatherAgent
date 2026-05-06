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
  | "jobTitle"
  | "managerName"
  | "managerEmail"
  | "startDate"
  | "workLocation"
  | "preferredCommunication"
  | "laptopOsPreference"
  | "accessNeeded"
  | "emergencyContactName"
  | "emergencyContactPhone"
  | "timeZone"
  | "accessibilityRequirements"
  | "ndaAcknowledged"
  | "securityTrainingAccepted"
  | "codeOfConductAccepted"
  | "dataPrivacyAcknowledged"
  | "deviceUsagePolicyAccepted"
  | "complete";

interface OnboardingSession {
  step: OnboardingStep;
  fullName?: string;
  email?: string;
  jobTitle?: string;
  managerName?: string;
  managerEmail?: string;
  startDate?: string;
  workLocation?: string;
  preferredCommunication?: string;
  laptopOsPreference?: string;
  accessNeeded?: string[];
  emergencyContactName?: string;
  emergencyContactPhone?: string;
  timeZone?: string;
  accessibilityRequirements?: string;
  ndaAcknowledged?: boolean;
  securityTrainingAccepted?: boolean;
  codeOfConductAccepted?: boolean;
  dataPrivacyAcknowledged?: boolean;
  deviceUsagePolicyAccepted?: boolean;
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

function isValidEmail(value: string) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function isValidPhone(value: string) {
  return /^[0-9()+\-\s]{7,20}$/.test(value);
}

function isValidDate(value: string) {
  return /^\d{4}-\d{2}-\d{2}$/.test(value) || /^\d{2}\/\d{2}\/\d{4}$/.test(value);
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
          { title: "Job Title", value: session.jobTitle ?? "-" },
          { title: "Manager", value: session.managerName ?? "-" },
          { title: "Manager Email", value: session.managerEmail ?? "-" },
          { title: "Start Date", value: session.startDate ?? "-" },
          { title: "Work Location", value: session.workLocation ?? "-" },
          {
            title: "Preferred Communication",
            value: session.preferredCommunication ?? "-",
          },
          { title: "Laptop OS", value: session.laptopOsPreference ?? "-" },
          {
            title: "Access Needed",
            value: session.accessNeeded?.length
              ? session.accessNeeded.join(", ")
              : "-",
          },
          {
            title: "Emergency Contact",
            value: session.emergencyContactName
              ? `${session.emergencyContactName} (${session.emergencyContactPhone ?? "-"})`
              : "-",
          },
          { title: "Time Zone", value: session.timeZone ?? "-" },
          {
            title: "Accessibility Requirements",
            value: session.accessibilityRequirements || "-",
          },
          {
            title: "NDA Acknowledged",
            value: session.ndaAcknowledged ? "Yes" : "No",
          },
          {
            title: "Security Training Accepted",
            value: session.securityTrainingAccepted ? "Yes" : "No",
          },
          {
            title: "Code of Conduct Accepted",
            value: session.codeOfConductAccepted ? "Yes" : "No",
          },
          {
            title: "Data Privacy Acknowledged",
            value: session.dataPrivacyAcknowledged ? "Yes" : "No",
          },
          {
            title: "Device Usage Policy Accepted",
            value: session.deviceUsagePolicyAccepted ? "Yes" : "No",
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

    const answer = (key: string) => submittedData[key] ?? text;

    if (session.step === "start") {
      await context.sendActivity("Great! Let’s begin onboarding. First question:");
      await sendCard(
        context,
        createQuestionCard(
          "Step 1 of 20: Employee Full Name",
          "What is your full name?",
          "fullName"
        )
      );
      session.step = "fullName";
      return;
    }

    if (session.step === "fullName") {
      const fullName = answer("fullName");
      if (!fullName) {
        await context.sendActivity("Please provide your full name.");
        return;
      }

      session.fullName = fullName;
      session.step = "email";
      await sendCard(
        context,
        createQuestionCard(
          "Step 2 of 20: Work Email",
          "What is your work email?",
          "email"
        )
      );
      return;
    }

    if (session.step === "email") {
      const email = answer("email");
      if (!email || !isValidEmail(email)) {
        await context.sendActivity("Please provide a valid email address.");
        return;
      }

      session.email = email;
      session.step = "jobTitle";
      await sendCard(
        context,
        createQuestionCard(
          "Step 3 of 20: Job Title",
          "What is your job title?",
          "jobTitle"
        )
      );
      return;
    }

    if (session.step === "jobTitle") {
      const jobTitle = answer("jobTitle");
      if (!jobTitle) {
        await context.sendActivity("Please provide your job title.");
        return;
      }

      session.jobTitle = jobTitle;
      session.step = "managerName";
      await sendCard(
        context,
        createQuestionCard(
          "Step 4 of 20: Manager Name",
          "Who is your manager?",
          "managerName"
        )
      );
      return;
    }

    if (session.step === "managerName") {
      const managerName = answer("managerName");
      if (!managerName) {
        await context.sendActivity("Please provide your manager's name.");
        return;
      }

      session.managerName = managerName;
      session.step = "managerEmail";
      await sendCard(
        context,
        createQuestionCard(
          "Step 5 of 20: Manager Email",
          "What is your manager's email?",
          "managerEmail"
        )
      );
      return;
    }

    if (session.step === "managerEmail") {
      const managerEmail = answer("managerEmail");
      if (!managerEmail || !isValidEmail(managerEmail)) {
        await context.sendActivity("Please provide a valid manager email address.");
        return;
      }

      session.managerEmail = managerEmail;
      session.step = "startDate";
      await sendCard(
        context,
        createQuestionCard(
          "Step 6 of 20: Start Date",
          "What is your planned start date? (YYYY-MM-DD or MM/DD/YYYY)",
          "startDate"
        )
      );
      return;
    }

    if (session.step === "startDate") {
      const startDate = answer("startDate");
      if (!startDate || !isValidDate(startDate)) {
        await context.sendActivity(
          "Please provide a valid date in YYYY-MM-DD or MM/DD/YYYY format."
        );
        return;
      }

      session.startDate = startDate;
      session.step = "workLocation";
      await sendCard(
        context,
        createQuestionCard(
          "Step 7 of 20: Work Location",
          "What is your work location?",
          "workLocation",
          ["Office", "Remote", "Hybrid"]
        )
      );
      return;
    }

    if (session.step === "workLocation") {
      const workLocation = answer("workLocation");
      if (!workLocation) {
        await context.sendActivity("Please choose a work location.");
        return;
      }

      session.workLocation = workLocation;
      session.step = "preferredCommunication";
      await sendCard(
        context,
        createQuestionCard(
          "Step 8 of 20: Preferred Communication",
          "Which communication tool do you prefer?",
          "preferredCommunication",
          ["Teams", "Email", "SMS"]
        )
      );
      return;
    }

    if (session.step === "preferredCommunication") {
      const preferredCommunication = answer("preferredCommunication");
      if (!preferredCommunication) {
        await context.sendActivity("Please choose your preferred communication tool.");
        return;
      }

      session.preferredCommunication = preferredCommunication;
      session.step = "laptopOsPreference";
      await sendCard(
        context,
        createQuestionCard(
          "Step 9 of 20: Laptop OS Preference",
          "Which laptop OS do you prefer?",
          "laptopOsPreference",
          ["Windows", "macOS"]
        )
      );
      return;
    }

    if (session.step === "laptopOsPreference") {
      const laptopOsPreference = answer("laptopOsPreference");
      if (!laptopOsPreference) {
        await context.sendActivity("Please choose a laptop OS preference.");
        return;
      }

      session.laptopOsPreference = laptopOsPreference;
      session.step = "accessNeeded";
      await sendCard(
        context,
        createQuestionCard(
          "Step 10 of 20: Access Needed",
          "Which of these systems do you need access to?",
          "accessNeeded",
          ["GitHub", "Jira", "CRM", "HR portal"],
          true
        )
      );
      return;
    }

    if (session.step === "accessNeeded") {
      const rawAccess = answer("accessNeeded");
      if (!rawAccess) {
        await context.sendActivity("Please select at least one access option.");
        return;
      }

      session.accessNeeded = rawAccess
        .split(",")
        .map((item) => item.trim())
        .filter(Boolean);
      session.step = "emergencyContactName";
      await sendCard(
        context,
        createQuestionCard(
          "Step 11 of 20: Emergency Contact Name",
          "What is the name of your emergency contact?",
          "emergencyContactName"
        )
      );
      return;
    }

    if (session.step === "emergencyContactName") {
      const emergencyContactName = answer("emergencyContactName");
      if (!emergencyContactName) {
        await context.sendActivity("Please provide an emergency contact name.");
        return;
      }

      session.emergencyContactName = emergencyContactName;
      session.step = "emergencyContactPhone";
      await sendCard(
        context,
        createQuestionCard(
          "Step 12 of 20: Emergency Contact Phone",
          "What is their phone number?",
          "emergencyContactPhone"
        )
      );
      return;
    }

    if (session.step === "emergencyContactPhone") {
      const emergencyContactPhone = answer("emergencyContactPhone");
      if (!emergencyContactPhone || !isValidPhone(emergencyContactPhone)) {
        await context.sendActivity(
          "Please provide a valid emergency contact phone number."
        );
        return;
      }

      session.emergencyContactPhone = emergencyContactPhone;
      session.step = "timeZone";
      await sendCard(
        context,
        createQuestionCard(
          "Step 13 of 20: Time Zone",
          "What time zone are you in?",
          "timeZone"
        )
      );
      return;
    }

    if (session.step === "timeZone") {
      const timeZone = answer("timeZone");
      if (!timeZone) {
        await context.sendActivity("Please provide your time zone.");
        return;
      }

      session.timeZone = timeZone;
      session.step = "accessibilityRequirements";
      await sendCard(
        context,
        createQuestionCard(
          "Step 14 of 20: Accessibility Requirements",
          "Do you have any accessibility requirements?",
          "accessibilityRequirements"
        )
      );
      return;
    }

    if (session.step === "accessibilityRequirements") {
      const accessibilityRequirements = answer("accessibilityRequirements");
      session.accessibilityRequirements = accessibilityRequirements || "None";
      session.step = "ndaAcknowledged";
      await sendCard(
        context,
        createQuestionCard(
          "Step 15 of 20: NDA Acknowledgement",
          "Do you acknowledge the NDA?",
          "ndaAcknowledged",
          ["Yes", "No"]
        )
      );
      return;
    }

    if (session.step === "ndaAcknowledged") {
      const ndaAnswer = answer("ndaAcknowledged").toLowerCase();
      if (ndaAnswer !== "yes" && ndaAnswer !== "no") {
        await context.sendActivity("Please answer Yes or No for NDA acknowledgement.");
        return;
      }

      session.ndaAcknowledged = ndaAnswer === "yes";
      session.step = "securityTrainingAccepted";
      await sendCard(
        context,
        createQuestionCard(
          "Step 16 of 20: Security Training",
          "Do you accept security training assignment?",
          "securityTrainingAccepted",
          ["Yes", "No"]
        )
      );
      return;
    }

    if (session.step === "securityTrainingAccepted") {
      const securityAnswer = answer("securityTrainingAccepted").toLowerCase();
      if (securityAnswer !== "yes" && securityAnswer !== "no") {
        await context.sendActivity(
          "Please answer Yes or No for security training acceptance."
        );
        return;
      }

      session.securityTrainingAccepted = securityAnswer === "yes";
      session.step = "codeOfConductAccepted";
      await sendCard(
        context,
        createQuestionCard(
          "Step 17 of 20: Code of Conduct",
          "Do you accept the code of conduct?",
          "codeOfConductAccepted",
          ["Yes", "No"]
        )
      );
      return;
    }

    if (session.step === "codeOfConductAccepted") {
      const conductAnswer = answer("codeOfConductAccepted").toLowerCase();
      if (conductAnswer !== "yes" && conductAnswer !== "no") {
        await context.sendActivity(
          "Please answer Yes or No for code of conduct acceptance."
        );
        return;
      }

      session.codeOfConductAccepted = conductAnswer === "yes";
      session.step = "dataPrivacyAcknowledged";
      await sendCard(
        context,
        createQuestionCard(
          "Step 18 of 20: Data Privacy",
          "Do you acknowledge the data privacy policy?",
          "dataPrivacyAcknowledged",
          ["Yes", "No"]
        )
      );
      return;
    }

    if (session.step === "dataPrivacyAcknowledged") {
      const privacyAnswer = answer("dataPrivacyAcknowledged").toLowerCase();
      if (privacyAnswer !== "yes" && privacyAnswer !== "no") {
        await context.sendActivity(
          "Please answer Yes or No for data privacy acknowledgement."
        );
        return;
      }

      session.dataPrivacyAcknowledged = privacyAnswer === "yes";
      session.step = "deviceUsagePolicyAccepted";
      await sendCard(
        context,
        createQuestionCard(
          "Step 19 of 20: Device Usage Policy",
          "Do you accept the device usage policy?",
          "deviceUsagePolicyAccepted",
          ["Yes", "No"]
        )
      );
      return;
    }

    if (session.step === "deviceUsagePolicyAccepted") {
      const devicePolicyAnswer = answer("deviceUsagePolicyAccepted").toLowerCase();
      if (devicePolicyAnswer !== "yes" && devicePolicyAnswer !== "no") {
        await context.sendActivity(
          "Please answer Yes or No for device usage policy acceptance."
        );
        return;
      }

      session.deviceUsagePolicyAccepted = devicePolicyAnswer === "yes";
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
