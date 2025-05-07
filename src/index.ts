import { Client } from "@microsoft/microsoft-graph-client";
import { ClientSecretCredential } from "@azure/identity";
import { Env } from './env';

export default {
	async fetch(request, env: Env, ctx): Promise<Response> {
		if (request.method === 'POST') {
			const response = await request.json() as {
				name: string;
				email: string;
				phone: string;
				message: string;
				'cf-turnstile-response': string;
			};
			const { name, email, phone, message } = response;

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

			if (!name || !email || !phone || !message) {
				return new Response(JSON.stringify({
					message: "Missing required fields"
				}), {
					status: 400,
					headers: {
						"Content-Type": "application/json",
					}
				});
			}

			try {
				const client = await getClient(env.MICROSOFT_GRAPH_CLIENT_ID, env.MICROSOFT_GRAPH_TENANT_ID, env.MICROSOFT_GRAPH_CLIENT_SECRET);
				await sendEmail(client, name, email, phone, message);
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

export async function sendEmail(client: Client, name: string, email_address: string, phone: string, message: string) {
	const email = {
		message: {
			subject: `Contact Form Submission from ${name}`,
			body: {
				contentType: "Text",
				content: `Name: ${name}\nEmail: ${email_address}\nPhone: ${phone}\nMessage: ${message}`
			},
			toRecipients: [
				{
					emailAddress: {
						address: "nathan@heathcotetech.com.au"
					}
				}
			]
		},
		saveToSentItems: false
	};

	await client.api('/users/nathan@heathcotetech.com.au/sendMail').post(email);
}
