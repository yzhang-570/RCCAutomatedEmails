/** @type {Object<string, {label: string, build: Function}>} */
const EMAIL_TEMPLATES = {
  STATUS_UPDATE: {
    label: "Status Update",
    build: buildStatusUpdateEmail
  },
  ACTIVE_CELEBRATION: {
    label: "Active Member Celebration",
    build: buildActiveMemberEmail
  },
  WELCOME: {
    label: "Welcome", // note: currently unused, welcome emails are never queued 
                      // (auto-sends without queue on form submission),
                      // but manually queueing a Welcome email is supported
    build: buildWelcomeEmail
  },
}



/**
 * Returns whether labelName matches a registered EMAIL_TEMPLATES entry. Pure - callers
 * are responsible for logging/handling an invalid label themselves.
 *
 * @param {string} labelName
 * @returns {boolean}
 */
function isValidEmailType_(labelName) {
  return Object.values(EMAIL_TEMPLATES).some(template =>
    normalize_(template.label) === normalize_(labelName));
}



/**
 * Builds the dynamic "Your RCC Status Update" email for a member, based on their current
 * Members DB stats. Includes the non-member registration prompt and/or ambassador
 * progress section only when applicable.
 *
 * @param {MemberRecord} record - a member's engagement record
 * @param {number} meetingsRequired - total # of AmbMeet type check-ins found
 * @param {number} eventName - event that triggered a status update
 * @returns {{subject: string, body: string}}
 */
function buildStatusUpdateEmail(record, meetingsRequired, eventName) {
  const firstName = getFirstName_(record.name) || "there";
  const activeMember = record.social >= 1 && record.nonSocial >= 1;

  const getActiveMemberStatus = () => {
    if (record.membershipStatus !== "Member") {
      return activeMember
        ? "{You're already qualified! Register for RCC Membership to activate your Active Member status:|gray}" + `${CONFIG.membershipFormLink}`
        : "{Become a member to be eligible to be an Active Member. Register for RCC Membership:|gray}" + `${CONFIG.membershipFormLink}`
    }
    return activeMember
      ? "{Active Member!|green}"
      : "{In Progress|orange}"
  }

  let body = `Hi ${firstName},

Thank you for your involvement with RCC! We saw that you recently attended ${eventName}. Here’s a quick update on your progress this semester.

**Active Member Progress**
Social event: ${record.social} / 1 fulfilled
Non-social event: ${record.nonSocial} / 1 fulfilled
Status: ${getActiveMemberStatus()}`;

  if (record.isAmbassador) {
    const meetingsCredited = record.ambMeetAttended + Math.min(CONFIG.maxExcusals, record.ambMeetExcused);
    const reqMet = activeMember && record.tabling >= 1 && meetingsCredited === meetingsRequired;
    body += `

**Ambassador Requirements**
Ambassador General Meetings: ${meetingsCredited} / ${meetingsRequired} fulfilled ${record.ambMeetExcused >= CONFIG.maxExcusals ? "{- Note: You have used your permitted one ambassador excusal. You have no remaining ambassador excusals this semester.|gray}" : ""}
Tabling: ${record.tabling} / 1 fulfilled
Active Member Requirement: ${activeMember ? "Completed" : "In Progress"}
Status: ${reqMet ? "{Fulfilled|green}" : "{In Progress|orange}"}`;
  }

  body += `
  
We hope to see you at our upcoming opportunities and events!

Best,
Responsible Computing Club
Email: ${CONFIG.socialLinks.email}
Discord: ${CONFIG.socialLinks.discord}
Instagram: ${CONFIG.socialLinks.instagram}
LinkedIn: ${CONFIG.socialLinks.linkedin}`;

  return { subject: "Your RCC Status Update", body };
}



/**
 * Builds the dynamic "You're an Active Member!" email for a member, based on their selected track.
 *
 * @param {string} fullName - a member's full name
 * @returns {{subject: string, body: string}}
 */
function buildActiveMemberEmail(fullName) {
  const body =
`Hi ${fullName},

Congratulations, you're an RCC Active Member of Fall 2026! We wanted to recognize you for your continued involvement with RCC.

**As an RCC Active Member this semester, you'll receive:**
- Priority consideration for company tours and other exclusive events
- The Active Member role in Discord
- Early notice for limited opportunities

By showing up and engaging, you've helped shape the RCC community, and we’re grateful to have you with us.

We hope that you'll continue to learn with, grow, and develop our community together!

Best,
Responsible Computing Club
Instagram: ${CONFIG.socialLinks.instagram}
Discord: ${CONFIG.socialLinks.discord}
LinkedIn: ${CONFIG.socialLinks.linkedin}
`;

  return { subject: "You're an Active Member!", body };
}




/**
 * Lane-specific welcome email builders, keyed by the exact "II. Choose Your Lane!" option
 * text from the Membership Form. Matched case/whitespace-insensitively via findLaneBuilder_.
 *
 * @type {Object<string, (fullName: string) => {subject: string, body: string}>}
 */
const WELCOME_LANE_BUILDERS = { // takes 1 keyword for resilience against name changes
  "Career": buildCareerIndustryWelcomeEmail,
  "Projects": buildProjectsWelcomeEmail,
  "Explore": buildLearnExploreWelcomeEmail,
  "Community": buildCommunityLeadershipWelcomeEmail
};

/**
 * Finds the lane builder matching `track`, case/whitespace-insensitively.
 *
 * @param {string} track
 * @returns {?((fullName: string) => {subject: string, body: string})}
 */
function findLaneBuilder_(track) {
  const normTrack = normalize_(track);
  const key = Object.keys(WELCOME_LANE_BUILDERS).find(lane => normTrack.includes(normalize_(lane)));
  return key ? WELCOME_LANE_BUILDERS[key] : null;
}

/**
 * Returns whether `track` matches one of the Membership Form's lane options. Pure -
 * callers are responsible for logging/handling an invalid lane themselves.
 *
 * @param {string} track
 * @returns {boolean}
 */
function isValidLane_(track) {
  return !!findLaneBuilder_(track);
}

/**
 * Builds the "Welcome to RCC!" email for a new registrant: lane-specific content when
 * `track` matches one of the Membership Form's lane options, otherwise the default template.
 *
 * @param {string} fullName - a member's full name
 * @param {string} track - a member's selected lane
 * @returns {{subject: string, body: string}}
 */
function buildWelcomeEmail(fullName, track) {
  const firstName = getFirstName_(fullName) || placeholderName;

  // send default only
  return buildDefaultWelcomeEmail_(firstName)

  // const laneBuilder = findLaneBuilder_(track);
  // return laneBuilder ? laneBuilder(firstName) : buildDefaultWelcomeEmail_(firstName);
}

/**
 * Wraps lane-specific content (intro + "How can I get involved?" bullets) with the
 * opening/closing shared by every lane welcome email.
 *
 * @param {string} firstName
 * @param {string} laneContent
 * @returns {{subject: string, body: string}}
 */
function buildLaneWelcomeEmail_(firstName, laneContent) {
  const body = `Hi ${firstName},

We’re excited to have you here! ${laneContent}

We’ve tailored this starting point to the interests you shared, but we welcome you to explore everything RCC has to offer. We hope to see you at our upcoming events!

Best,
Responsible Computing Club
Instagram: ${CONFIG.socialLinks.instagram}
Discord: ${CONFIG.socialLinks.discord}
LinkedIn: ${CONFIG.socialLinks.linkedin}
`;
  return { subject: "Welcome to RCC! 👋🎊", body };
}

/**
 * @param {string} firstName
 * @returns {{subject: string, body: string}}
 */
function buildCareerIndustryWelcomeEmail(firstName) {
  return buildLaneWelcomeEmail_(firstName,

`Whether you’re ready to build professional connections, explore career paths, or prepare for opportunities in industry, you’re in the right place.

**How can I get involved?**
- Company tours and events
- Networking with students and professionals
- Industry and speaker panels
- Professional development workshops and case competitions
- Internship and career opportunities`);
}

// Whether you’re looking to join your first project or strengthen your portfolio, RCC will connect you with opportunities to learn, contribute, and create.

/**
 * @param {string} firstName
 * @returns {{subject: string, body: string}}
 */
function buildProjectsWelcomeEmail(firstName) {
  return buildLaneWelcomeEmail_(firstName, 

`Whether you’re ready to turn ideas into action, join your first project, or strengthen your portfolio, you’re in the right place. 

**How can I get involved?**
- Consulting and student-led projects
- RCC project development program
- Opportunities to solve real-world problems
- Team-based technical and interdisciplinary work`);
}

// RCC will connect you with opportunities to explore new ideas, learn technical skills, and hear from different perspectives. 
/**
 * @param {string} firstName
 * @returns {{subject: string, body: string}}
 */
function buildLearnExploreWelcomeEmail(firstName) {
  return buildLaneWelcomeEmail_(firstName, 
    
`Whether you’re curious about emerging technologies, exploring their impact, or hearing from different perspectives, you’re in the right place.

**How can I get involved?**
- Technical workshops
- Research and exploration opportunities
- Interdisciplinary discussions about ethics and responsible innovation
- Faculty, student, and industry panels`);
}

// You’ll find opportunities to contribute to the community, develop your leadership skills, and help shape the RCC experience. 
/**
 * @param {string} firstName
 * @returns {{subject: string, body: string}}
 */
function buildCommunityLeadershipWelcomeEmail(firstName) {
  return buildLaneWelcomeEmail_(firstName, 
    
`Whether you’re ready to meet new people, get involved in leadership, or make an impact, you’re in the right place. 

**How can I get involved?**
- RCC ambassador (leadership) program
- Volunteering and event involvement
- Social and community-building events
- Opportunities to help shape future RCC programs`);
}

/**
 * Default "Welcome to RCC!" email, used when no lane is selected/recognized.
 *
 * @param {string} firstName
 * @returns {{subject: string, body: string}}
 */
function buildDefaultWelcomeEmail_(firstName) {

  const body =

`Hi ${firstName},

Welcome to the Responsible Computing Club (RCC)! We’re excited to have you in our community. RCC offers many opportunities to gain hands-on experience, meet new people, and grow professionally.

**How can I get involved?**
- Leadership opportunities
- Consulting and student-led projects
- Industry events, speaker panels, and workshops
- Networking and socials
- ...and more!

Stay on the lookout for our weekly RCC newsletters with upcoming events, opportunities, applications, and important updates. You can also stay up to date throughout the week on our Instagram and Discord.

**Happening Now** 👀
- 💡Ambassador Applications are open! Join or lead an RCC team and help design, build, and run our initiatives → [Learn more & apply here](${CONFIG.ambassadorApplicationLink})

We hope to see you at our upcoming RCC events!

Best,
Responsible Computing Club
Instagram: ${CONFIG.socialLinks.instagram}
Discord: ${CONFIG.socialLinks.discord}
LinkedIn: ${CONFIG.socialLinks.linkedin}`;

  return {subject: "Welcome to RCC! 👋🎊", body};
}
