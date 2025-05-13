import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";
import { Env } from './env';

export default {
	async fetch(request, env: Env, ctx): Promise<Response> {
		const origin = request.headers.get('origin');
		if (origin && origin !== 'https://infinitysupport.heathcotetech.com.au') {
			return new Response(
                JSON.stringify({message: "Invalid Origin"}),
                {
                    status: 403,
                    headers: {"Content-Type": "application/json"},
                }
            );
		}

		if (request.method === 'POST') {
			interface participantDetails {
				name: string;
				email: string;
				phone: string;
				dob: string;
				disability: string;
				behaviour: string;
			};

			interface services {
				support: boolean;
				community: boolean;
				allied: boolean;
				accomodation: boolean;
			}

			interface coordinator {
				name: string;
				email: string;
				phone: string;
				company: string;
			}

			interface plan {
				name: string;
				email: string;
				type: string;
			}

			interface ndis {
				ndisNumber: string;
				startDate: string;
				endDate: string;
			}

			interface days {
				monday: boolean;
				tuesday: boolean;
				wednesday: boolean;
				thursday: boolean;
				friday: boolean;
				saturday: boolean;
				sunday: boolean;
			}

			interface attachment {
				name: string;
				contentType: string;
				contentBytes: string;
			}

			const response = await request.json() as {
				name: string;
				email: string;
				phone: string;
				message:string;
				participant: participantDetails;
				services: services;
				coordinator: coordinator;
				plan: plan;
				ndis: ndis;
				days: days;
				attachments: attachment[];
				type: string;
				'cf-turnstile-response': string;
			};

			const token = response['cf-turnstile-response'];
    		const remote_ip = request.headers.get('cf-connecting-ip');
			const SECRET_KEY = env.TURNSTILE_SECRET;
			const url = 'https://challenges.cloudflare.com/turnstile/v0/siteverify';
			const result = await fetch(url, {
				method: 'POST',
				headers: {
					'Content-Type': 'application/json'
				},
				body: JSON.stringify({
					secret: SECRET_KEY,
					response: token,
					remoteip: remote_ip
				})
			});

			const cf_response: { success: boolean } = await result.json();
			if (!cf_response.success) {
				return new Response(JSON.stringify({
					message: "Human Verification has failed"
				}), {
					status: 403,
					headers: {
						"Content-Type": "application/json",
					}
				});
			}


			const { name, email, phone, message, participant, services, coordinator, plan, ndis, days, type } = response;
			if ((type === "contact" && (!name || !email || !phone || !message))) {
				return new Response(JSON.stringify({
					message: "Missing required fields"
				}), {
					status: 400,
					headers: {
						"Content-Type": "application/json",
					}
				});
			}

			let subject: string = "";
			let body: string = "";
			if (type === "contact") {
				subject = `Contact Form Submission from ${name}`;
				body = `Name: ${name}\nEmail: ${email}\nPhone: ${phone}\nMessage: ${message}`;
			} else if (type === "referral") {
				const { name, email, phone, dob, disability, behaviour } = participant;
				const { support, community, allied, accomodation } = services;
				const { name: coordinatorName, email: coordinatorEmail, phone: coordinatorPhone, company } = coordinator;
				const { name: planName, email: planEmail, type: planType } = plan;
				const plantype = plan.type;
				const { ndisNumber, startDate, endDate } = ndis;
				const { monday, tuesday, wednesday, thursday, friday, saturday, sunday } = days;

				subject = `Referral Form Submission for ${name}`;
				
				const serviceLabels: Record<string, string> = {
					support: "Support Coordination",
					community: "Community Access",
					allied: "Allied Health Assistants",
					accomodation: "Accomodation",
				}
				
				const requiredServices = Object.entries({support, community, allied, accomodation})
					.filter(([_key, value]) => value === true)
					.map(([key]) => serviceLabels[key])
					.join(", ");
				
				const daysOfWeek = Object.entries({monday, tuesday, wednesday, thursday, friday, saturday, sunday})
					.filter(([_key, value]) => value === true)
					.map(([key]) => key.charAt(0).toUpperCase() + key.slice(1))
					.join(", ");

				let planTypeString = "";
				if (plantype === "ndia") {
					planTypeString = "NDIA";
				} else if (plantype === "self-managed") {
					planTypeString = "Self Managed";
				} else if (plantype === "plan-managed") {
					planTypeString = "Plan Managed";
				}

				const participantDetailsString = `Name: ${name}\nEmail: ${email}\nPhone: ${phone}\nDate of Birth: ${dob}\nPrimary Disability: ${disability}\nPotential Risks/Behaviour Concerns: ${behaviour}`;
				const servicesDetails = `Services Requested: ${requiredServices}`;
				const coordinatorDetails = `Name: ${coordinatorName}\nEmail: ${coordinatorEmail}\nPhone: ${coordinatorPhone}\nCompany: ${company}`;
				const planDetails = `Name: ${planName}\nEmail: ${planEmail}\nPlan Type: ${planTypeString}`;
				const ndisDetails = `NDIS Number: ${ndisNumber}\nStart Date: ${startDate}\nEnd Date: ${endDate}`;

				body = `Participant Details:\n${participantDetailsString}\n\nServices Details:\n${servicesDetails}\n\nCoordinator Details:\n${coordinatorDetails}\n\nPlan Manager Details:\n${planDetails}\n\nNDIS Details:\n${ndisDetails}\n\nPreferred Support Days:\n${daysOfWeek}`;
			} else {
				return new Response(JSON.stringify({
					message: `Invalid form type: ${type}`
				}), {
					status: 400,
					headers: {
						"Content-Type": "application/json",
					}
				});
			}

			const fileAttachments = [];
			for (const attachment of response.attachments) {
				fileAttachments.push({
					...attachment,
					"@odata.type": "#microsoft.graph.fileAttachment",
				});
			}

			try {
				const client = await getClient(env.MICROSOFT_GRAPH_CLIENT_ID, env.MICROSOFT_GRAPH_TENANT_ID, env.MICROSOFT_GRAPH_CLIENT_SECRET);
				await sendEmail(client, subject, body, fileAttachments);
			} catch (error) {
				return new Response(JSON.stringify({
					message: `Unable to send an email...\n${error}`
				}), {
					status: 500,
					headers: {
						"Content-Type": "application/json",
					}
				});
			}

			return new Response(JSON.stringify({
				message: "Email has been sent"
			}), {
				status: 200,
				headers: {
					"Content-Type": "application/json",
				}
			});
		}

		return new Response(null, {
			status: 302,
			headers: {
			  Location: 'https://infinitysupport.heathcotetech.com.au/'
			}
		  });
	},
} satisfies ExportedHandler<Env>;

async function getClient(MICROSOFT_GRAPH_CLIENT_ID: string, MICROSOFT_GRAPH_TENANT_ID: string, MICROSOFT_GRAPH_CLIENT_SECRET: string) {
    if (!MICROSOFT_GRAPH_CLIENT_ID || !MICROSOFT_GRAPH_TENANT_ID || !MICROSOFT_GRAPH_CLIENT_SECRET) {
        throw new Error("Missing Microsoft Graph credentials");
    }
    
    const credentials = new ClientSecretCredential(
        MICROSOFT_GRAPH_TENANT_ID,
        MICROSOFT_GRAPH_CLIENT_ID,
        MICROSOFT_GRAPH_CLIENT_SECRET
    );
    
    const token = await credentials.getToken("https://graph.microsoft.com/.default");
    const client = Client.init({
        authProvider: (done) => {
            done(null, token?.token ?? "")
        }
    });
    
    return client;
}

export async function sendEmail(client: Client, subject: string, message: string, fileAttachments: { "@odata.type": string, name: string; contentType: string; contentBytes: string }[] | null) {
	const email = {
		message: {
			subject: subject,
			body: {
				contentType: "Text",
				content: message
			},
			toRecipients: [
				{
					emailAddress: {
						address: "nathan@heathcotetech.com.au"
					}
				}
			],
			...(fileAttachments !== null && fileAttachments.length > 0 ? { attachments: fileAttachments } : {}),
		},
		saveToSentItems: false
	};

	await client.api('/users/nathan@heathcotetech.com.au/sendMail').post(email);
}
