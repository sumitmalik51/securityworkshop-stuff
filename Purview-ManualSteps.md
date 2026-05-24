# Purview - Remaining Manual Steps

These steps require the Microsoft Purview Portal UI and cannot be automated via PowerShell.

---

## 1. DLP – Turn on Analytics Recommendation

1. Open **InPrivate** browser → [Microsoft Purview Portal](https://purview.microsoft.com)
2. Navigate to **Solutions → DLP → Overview**
3. Turn on: **"Turn on analytics for risk detection and policy refinement opportunities (Preview)"**

---

## 2. DLP – Deploy Triage Agent

1. Navigate to **Agents** (left nav) → **Explore agents**
2. Under **Triage agent in Data Loss Prevention** → click **View details**
3. Click **Set up** (bottom-right)
4. Trigger: select **"Automatic when a new alert is detected (recommended)"**
5. Alert timeframe: select **"Last 30 days (recommended)"**
6. Click **Start**
7. Confirm message: **"Agent is now active"**
8. Click **View results**

---

## 3. DLP – Customize Triage Agent

1. On the Alerts page, switch toggle from **Standard** to **Triage Agent**
2. Click **Customize Agent**
3. Expand **Agent configuration** → click **Edit** under "Set agent scope"
4. **Select all DLP policies** → click **Select policies**
5. Expand **Custom instruction (Preview)**
6. Enter instruction:
   > Focus on alerts with content that is tax or finance related and contains more than five credit card numbers or SSNs
7. Click **Generate** / **Interpretation** to review
8. Click **Save**

---

## 4. IRM – Deploy Triage Agent

1. Navigate to **Agents** (left nav) → **Explore agents**
2. Under **Triage agent in Insider Risk Management** → click **View details**
3. Click **Set up**
4. Trigger: select **"Manual on one alert at a time"**
5. Click **Start**
6. Confirm message: **"Agent is now active"**
7. Click **View results**

---

## 5. IRM – Customize Triage Agent

1. On Alerts page, switch toggle from **Standard** to **Triage Agent**
2. Click **Customize agent**
3. Expand **Agent configuration** → click **Edit** under "Set agent scope"
4. Select the required IRM policies
5. Click **Select policies**
6. Click **Save**

---

## 6. DSPM – Enable & Configure Posture Agent

Follow the lab guide: **Exercise 3, Tasks 1 and 2** completely.

---

## 7. DSI – Data Security Investigations

Follow the lab guide: **Exercise 4, Task 1** completely.

---

## Notes

- **Yash** will run the script to trigger DLP alerts in categorized form for all environments.
- **Yash** will run the script to trigger IRM alerts for all 200 users.
- Run `Purview-AutoSetup.ps1` first to ensure all role assignments and policies are in place before performing these manual steps.
