interface contact_response {
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
};

interface coordinator {
	name: string;
	email: string;
	phone: string;
	company: string;
};

interface plan {
	name: string;
	email: string;
	type: string;
};

interface ndis {
	ndisNumber: string;
	startDate: string;
	endDate: string;
};

interface days {
	monday: boolean;
	tuesday: boolean;
	wednesday: boolean;
	thursday: boolean;
	friday: boolean;
	saturday: boolean;
	sunday: boolean;
};

interface attachment {
	name: string;
	contentType: string;
	contentBytes: string;
};