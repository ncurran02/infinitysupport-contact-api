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
			const response = await request.json() as {
				name: string;
				email: string;
				phone: string;
				dob: string;
				disability: string;
				behaviour: string;
				support: string;
				community: string;
				allied: string;
				accomodation: string;
				message: string;
				type: string;
				attachment: {
					name: string;
					contentType: string;
					contentBytes: string;
				};
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


			const { name, email, phone, dob, disability, behaviour, support, community, allied, accomodation, message, type } = response;
			if ((type === "contact" && (!name || !email || !phone || !message)) || (type === "referral" && (!name || !email || !phone || !dob || !disability || !behaviour))) {
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
				subject = `Referral Form Submission for ${name}`;
				
				const serviceLabels: Record<string, string> = {
					support: "Support Coordination",
					community: "Community Access",
					allied: "Allied Health Assistants",
					accomodation: "Accomodation",
				}
				const requiredServices = [support, community, allied, accomodation].filter(service => service === "on").map(service => serviceLabels[service]).join(",");
				
				body = `Name: ${name}\nEmail: ${email}\nPhone: ${phone}\nDate of Birth: ${dob}\nPrimary Disability: ${disability}\nPotential Risks/Behaviour Concerns: ${behaviour}\nServices Requested: ${requiredServices}`;
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

			const attachment = response.attachment ? {
				...response.attachment,
				"@odata.type": "#microsoft.graph.fileAttachment",
			} : null;

			try {
				const client = await getClient(env.MICROSOFT_GRAPH_CLIENT_ID, env.MICROSOFT_GRAPH_TENANT_ID, env.MICROSOFT_GRAPH_CLIENT_SECRET);
				await sendEmail(client, subject, body, attachment);
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

export async function sendEmail(client: Client, subject: string, message: string, attachment: { "@odata.type": string, name: string; contentType: string; contentBytes: string } | null) {
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
			...(attachment !== null ? { attachments: [attachment] } : {}),
		},
		saveToSentItems: false
	};

	await client.api('/users/nathan@heathcotetech.com.au/sendMail').post(email);
}
