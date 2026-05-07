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
    await sendCard(context, {
      type: "AdaptiveCard",
      version: "1.5",
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      body: [
        {
          type: "Container",
          style: "emphasis",
          bleed: true,
          items: [
            {
              type: "TextBlock",
              text: "👋 Welcome to Mazik Global",
              weight: "Bolder",
              size: "Large",
              wrap: true,
            },
            {
              type: "TextBlock",
              text: "We're excited to have you with us. Let's help you get started with a smooth onboarding experience.",
              wrap: true,
              spacing: "Small",
            },
          ],
        },
        {
          type: "TextBlock",
          text: "About Mazik Global",
          weight: "Bolder",
          size: "Medium",
          spacing: "Medium",
        },
        {
          type: "TextBlock",
          text: "Mazik Global is a people-first technology consulting company focused on delivering enterprise innovation and measurable business outcomes.",
          wrap: true,
          isSubtle: true,
          spacing: "Small",
        },
        {
          type: "FactSet",
          spacing: "Medium",
          facts: [
            {
              title: "Company",
              value: "Mazik Global",
            },
            {
              title: "Core Values",
              value: "Innovation, Collaboration, Excellence",
            },
            {
              title: "Your First Step",
              value: "Complete your onboarding profile",
            },
          ],
        },
        {
          type: "TextBlock",
          text: "Who to Contact",
          weight: "Bolder",
          size: "Medium",
          spacing: "Large",
        },
        {
          type: "Container",
          style: "default",
          spacing: "Small",
          items: [
            {
              type: "TextBlock",
              text: "Leadership & Strategic Management",
              weight: "Bolder",
              wrap: true,
            },
            {
              type: "TextBlock",
              text: "• Faima Noor (Chief of Staff): Strategic HR operations and high-level internal coordination across Pakistan\n• Afzal Syed (Partner): Strategic partnership and business oversight\n• Muhammad Khan (Senior Finance Manager): Financial operations and administrative matters\n• Waqas Anwer (Director, Technical Services - FinOps): Technical delivery and financial operations management",
              wrap: true,
              spacing: "Small",
              isSubtle: true,
            },
          ],
        },
        {
          type: "Container",
          style: "default",
          spacing: "Medium",
          items: [
            {
              type: "TextBlock",
              text: "HR, Recruitment & Careers",
              weight: "Bolder",
              wrap: true,
            },
            {
              type: "TextBlock",
              text: "• Saba Umer (HR Manager): Primary contact for HR and talent management in Karachi\n• Mansoor Hussain Laghari (HR Executive | Talent Acquisition): Recruitment, hiring, and employee onboarding",
              wrap: true,
              spacing: "Small",
              isSubtle: true,
            },
          ],
        },
        {
          type: "Container",
          style: "default",
          spacing: "Medium",
          items: [
            {
              type: "TextBlock",
              text: "Technical Delivery & Solutions",
              weight: "Bolder",
              wrap: true,
            },
            {
              type: "TextBlock",
              text: "• Huzaifa Jalali (Managing Consultant): Leads consulting teams for complex enterprise solutions\n• Munam Ahmed (D365 Service Delivery Coordinator): Coordinates Dynamics 365 delivery and support\n• Haseeb Ahmed (Senior Solutions Engineer): Engineers technical solutions for client business needs\n• Munib Qazi (Technical Team Lead): Senior software engineer leading development teams",
              wrap: true,
              spacing: "Small",
              isSubtle: true,
            },
          ],
        },
        {
          type: "Container",
          style: "default",
          spacing: "Medium",
          items: [
            {
              type: "TextBlock",
              text: "Innovation & Specialist Teams",
              weight: "Bolder",
              wrap: true,
            },
            {
              type: "TextBlock",
              text: "• Team Mazikians: Specialists representing Mazik Global at tech summits (Shayan Zubair, Safa Rizwan, Hafsa Farooqui, Shehzed Akmal)\n• Anas Uddin: Technical specialist in Azure and Dynamics 365 CRM",
              wrap: true,
              spacing: "Small",
              isSubtle: true,
            },
          ],
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Start Onboarding",
          style: "positive",
          data: {
            action: "startOnboarding",
          },
        },
        {
          type: "Action.OpenUrl",
          title: "Employee Handbook",
          url: "https://mazikglobalinc.sharepoint.com/:b:/s/MazikGlobal-KarachiOffice2/IQAad2IxqJZnR7VkGwEntvEuAUK6Bw7C5VfvfCjfLfYgess?e=rlpCgb",
        },
      ],
    });
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
        type: "TextBlock",
        text: "Your onboarding process is completed. This information will be sent to HR right away for final processing.",
        wrap: true,
        spacing: "Small",
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
    actions: [
      {
        type: "Action.Submit",
        title: "Accept & Submit",
        style: "positive",
        data: {
          action: "acceptOnboarding",
        },
      },
      {
        type: "Action.Submit",
        title: "Start Onboarding Again",
        data: {
          action: "restartOnboarding",
        },
      },
    ],
  };
}

function createThankYouCard() {
  return {
    type: "AdaptiveCard",
    version: "1.5",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    body: [
      {
        type: "TextBlock",
        text: "Thank You! 🎉",
        weight: "Bolder",
        size: "Large",
        wrap: true,
      },
      {
        type: "TextBlock",
        text: "Your information has been confirmed and submitted successfully.",
        wrap: true,
        spacing: "Small",
      },
      {
        type: "TextBlock",
        text: "Congratulations on joining Mazik Global. We wish you great success and an amazing journey ahead!",
        wrap: true,
        spacing: "Small",
        isSubtle: true,
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
    const action = submittedData.action;

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

    if (action === "acceptOnboarding") {
      await sendCard(context, createThankYouCard());
      return;
    }

    if (action === "restartOnboarding") {
      onboardingSessions.set(conversationId, { step: "start" });
      await context.sendActivity("Onboarding restarted.");
      await context.sendActivity("Great! Let’s begin onboarding. First question:");
      await sendCard(
        context,
        createQuestionCard(
          "Step 1 of 20: Employee Full Name",
          "What is your full name?",
          "fullName"
        )
      );
      const restartedSession = getSession(conversationId);
      restartedSession.step = "fullName";
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
      return;
    }

    if (session.step === "complete") {
      if (text.toLowerCase() === "restart") {
        onboardingSessions.set(conversationId, { step: "start" });
        await context.sendActivity("Onboarding restarted.");
      } else {
        await context.sendActivity(
          "Onboarding is complete. Please use the card buttons to Accept & Submit or Start Onboarding Again."
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
