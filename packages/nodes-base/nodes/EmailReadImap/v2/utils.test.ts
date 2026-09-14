import { type ImapSimple } from '@n8n/imap';
import { mock, mockDeep } from 'vitest-mock-extended';
import { returnJsonArray } from 'n8n-core';
import type { INode, ITriggerFunctions, IDataObject } from 'n8n-workflow';

import { getNewEmails } from './utils';

describe('Test IMap V2 utils', () => {
	afterEach(() => vi.resetAllMocks());

	describe('getNewEmails', () => {
		const triggerFunctions = mockDeep<ITriggerFunctions>({
			helpers: { returnJsonArray },
		});

		const message = {
			attributes: {
				uuid: 1,
				uid: 873,
				struct: {},
			},
			parts: [
				{ which: '', body: 'Body content' },
				{ which: 'HEADER', body: 'h' },
				{ which: 'TEXT', body: 'txt' },
			],
		};

		const connectionReturning = (...batches: unknown[][]) => {
			const search = vi.fn();
			for (const batch of batches) search.mockResolvedValueOnce(batch);
			return mock<ImapSimple>({
				search,
				downloadText: vi.fn().mockResolvedValue('text'),
				downloadAttachments: vi.fn().mockResolvedValue([]),
			});
		};

		it.each([
			{
				format: 'resolved',
				expected: {
					json: {
						attachments: undefined,
						headers: { '': 'Body content' },
						headerLines: undefined,
						html: false,
						attributes: { uid: 873 },
					},
					binary: undefined,
				},
			},
			{
				format: 'simple',
				expected: {
					json: {
						textHtml: 'text',
						textPlain: 'text',
						metadata: { '0': 'h' },
						attributes: { uid: 873 },
					},
				},
			},
			{
				format: 'raw',
				expected: { json: { raw: 'txt' } },
			},
		])('returns a new email in $format format', async ({ format, expected }) => {
			triggerFunctions.getNode.mockReturnValue(mock<INode>({ typeVersion: 2.1 }));
			triggerFunctions.getNodeParameter.calledWith('format').mockReturnValue(format);
			triggerFunctions.getNodeParameter
				.calledWith('dataPropertyAttachmentsPrefixName')
				.mockReturnValue('resolved');
			triggerFunctions.getNodeParameter.calledWith('downloadAttachments').mockReturnValue(false);
			triggerFunctions.getWorkflowStaticData.mockReturnValue({});

			const onEmailBatch = vi.fn();
			await getNewEmails.call(triggerFunctions, {
				imapConnection: connectionReturning([message]),
				searchCriteria: [],
				postProcessAction: '',
				onEmailBatch,
			});

			expect(onEmailBatch).toHaveBeenCalledTimes(1);
			expect(onEmailBatch).toHaveBeenCalledWith([expected]);
		});

		const rawEmailWithDate = 'Date: Wed, 01 Jan 2020 12:00:00 +0000\r\nSubject: Hello\r\n\r\nBody';

		const resolveEmailWithDate = async (typeVersion: number) => {
			const messageWithDate = {
				attributes: { uuid: 1, uid: 950, struct: {} },
				parts: [{ which: '', body: rawEmailWithDate }],
			};
			const localConnection = connectionReturning([messageWithDate]);

			triggerFunctions.getNode.mockReturnValue(mock<INode>({ typeVersion }));
			triggerFunctions.getNodeParameter.calledWith('format').mockReturnValue('resolved');
			triggerFunctions.getNodeParameter
				.calledWith('dataPropertyAttachmentsPrefixName')
				.mockReturnValue('resolved');
			triggerFunctions.getWorkflowStaticData.mockReturnValue({});

			const onEmailBatch = vi.fn();
			await getNewEmails.call(triggerFunctions, {
				imapConnection: localConnection,
				searchCriteria: [],
				postProcessAction: '',
				onEmailBatch,
			});
			return onEmailBatch.mock.calls[0][0][0] as { json: { date: unknown } };
		};

		it('should return the mail date as an ISO string in resolved format on version 2.2', async () => {
			const item = await resolveEmailWithDate(2.2);
			expect(item.json.date).toBe('2020-01-01T12:00:00.000Z');
		});

		it('should keep the mail date as a Date object in resolved format before version 2.2', async () => {
			const item = await resolveEmailWithDate(2.1);
			expect(item.json.date).toBeInstanceOf(Date);
		});

		it('should skip emails with missing HEADER part in simple format', async () => {
			const messageWithoutHeader = {
				attributes: { uuid: 1, uid: 900, struct: {} },
				parts: [{ which: 'TEXT', body: 'txt' }],
			};

			const localConnection = connectionReturning([messageWithoutHeader]);

			triggerFunctions.getNode.mockReturnValue(mock<INode>({ typeVersion: 2.1 }));
			triggerFunctions.getNodeParameter.calledWith('format').mockReturnValue('simple');
			triggerFunctions.getNodeParameter.calledWith('downloadAttachments').mockReturnValue(false);
			triggerFunctions.getWorkflowStaticData.mockReturnValue({});

			const onEmailBatch = vi.fn();
			await getNewEmails.call(triggerFunctions, {
				imapConnection: localConnection,
				searchCriteria: [],
				postProcessAction: 'nothing',
				onEmailBatch,
			});

			expect(onEmailBatch).toHaveBeenCalledWith([]);
			expect(triggerFunctions.logger.warn).toHaveBeenCalledWith(
				expect.stringContaining('HEADER part missing'),
			);
		});

		it('should only mark processed emails as read, not filtered-out ones', async () => {
			const staticData: IDataObject = { lastMessageUid: 873 };
			const messages = [
				{
					attributes: { uuid: 1, uid: 870, struct: {} },
					parts: [
						{ which: 'TEXT', body: 'txt' },
						{ which: 'HEADER', body: { from: ['a@b.com'] } },
					],
				},
				{
					attributes: { uuid: 2, uid: 873, struct: {} },
					parts: [
						{ which: 'TEXT', body: 'txt' },
						{ which: 'HEADER', body: { from: ['b@b.com'] } },
					],
				},
				{
					attributes: { uuid: 3, uid: 875, struct: {} },
					parts: [
						{ which: 'TEXT', body: 'txt' },
						{ which: 'HEADER', body: { from: ['c@b.com'] } },
					],
				},
			];

			const localConnection = connectionReturning(messages);

			triggerFunctions.getNode.mockReturnValue(mock<INode>({ typeVersion: 2.1 }));
			triggerFunctions.getNodeParameter.calledWith('format').mockReturnValue('simple');
			triggerFunctions.getNodeParameter.calledWith('downloadAttachments').mockReturnValue(false);
			triggerFunctions.getWorkflowStaticData.mockReturnValue(staticData);

			const onEmailBatch = vi.fn();
			await getNewEmails.call(triggerFunctions, {
				imapConnection: localConnection,
				searchCriteria: [],
				postProcessAction: 'read',
				onEmailBatch,
			});

			expect(localConnection.addFlags).toHaveBeenCalledTimes(1);
			expect(localConnection.addFlags).toHaveBeenCalledWith([875], '\\SEEN');
			expect(onEmailBatch).toHaveBeenCalledWith([
				expect.objectContaining({
					json: expect.objectContaining({
						attributes: { uid: 875 },
					}),
				}),
			]);
		});

		it('should not call addFlags when all messages are filtered out', async () => {
			const staticData: IDataObject = { lastMessageUid: 873 };
			const messages = [
				{
					attributes: { uuid: 1, uid: 873, struct: {} },
					parts: [
						{ which: 'TEXT', body: 'txt' },
						{ which: 'HEADER', body: { from: ['a@b.com'] } },
					],
				},
			];

			const localConnection = connectionReturning(messages);

			triggerFunctions.getNode.mockReturnValue(mock<INode>({ typeVersion: 2.1 }));
			triggerFunctions.getNodeParameter.calledWith('format').mockReturnValue('simple');
			triggerFunctions.getNodeParameter.calledWith('downloadAttachments').mockReturnValue(false);
			triggerFunctions.getWorkflowStaticData.mockReturnValue(staticData);

			const onEmailBatch = vi.fn();
			await getNewEmails.call(triggerFunctions, {
				imapConnection: localConnection,
				searchCriteria: [],
				postProcessAction: 'read',
				onEmailBatch,
			});

			expect(localConnection.addFlags).not.toHaveBeenCalled();
			expect(onEmailBatch).toHaveBeenCalledWith([]);
		});

		it('should update lastMessageUid between batch iterations to prevent duplicates', async () => {
			const staticData: IDataObject = {};
			const batch1 = Array.from({ length: 20 }, (_, i) => ({
				attributes: { uuid: i + 1, uid: i + 1, struct: {} },
				parts: [
					{ which: 'TEXT', body: 'txt' },
					{ which: 'HEADER', body: { from: [`user${i}@test.com`] } },
				],
			}));
			const batch2 = [
				{
					attributes: { uuid: 20, uid: 20, struct: {} },
					parts: [
						{ which: 'TEXT', body: 'txt' },
						{ which: 'HEADER', body: { from: ['user20@test.com'] } },
					],
				},
				{
					attributes: { uuid: 21, uid: 21, struct: {} },
					parts: [
						{ which: 'TEXT', body: 'txt' },
						{ which: 'HEADER', body: { from: ['user21@test.com'] } },
					],
				},
			];

			const localConnection = connectionReturning(batch1, batch2);

			triggerFunctions.getNode.mockReturnValue(mock<INode>({ typeVersion: 2.1 }));
			triggerFunctions.getNodeParameter.calledWith('format').mockReturnValue('simple');
			triggerFunctions.getNodeParameter.calledWith('downloadAttachments').mockReturnValue(false);
			triggerFunctions.getWorkflowStaticData.mockReturnValue(staticData);

			const allBatchedUids: number[][] = [];
			const onEmailBatch = vi.fn().mockImplementation((data) => {
				allBatchedUids.push(
					data.map((item: { json: { attributes: { uid: number } } }) => item.json.attributes.uid),
				);
			});

			await getNewEmails.call(triggerFunctions, {
				imapConnection: localConnection,
				searchCriteria: [],
				postProcessAction: 'nothing',
				onEmailBatch,
			});

			expect(onEmailBatch).toHaveBeenCalledTimes(2);

			const allEmittedUids = allBatchedUids.flat();
			const uniqueUids = new Set(allEmittedUids);
			expect(allEmittedUids.length).toBe(uniqueUids.size);

			expect(allBatchedUids[0]).toHaveLength(20);
			expect(allBatchedUids[1]).toEqual([21]);
			expect(staticData.lastMessageUid).toBe(21);
		});

		/**
		 * A message that produces no item was never handed to the workflow. The
		 * watermark is what every later search filters on, so if it moves past
		 * that message, the message is gone for good: no execution, no error, no
		 * entry anywhere in the UI (#36681).
		 */
		describe('the watermark and a message that made no item', () => {
			const withHeader = (uid: number) => ({
				attributes: { uuid: uid, uid, struct: {} },
				parts: [
					{ which: 'TEXT', body: 'txt' },
					{ which: 'HEADER', body: { from: [`user${uid}@test.com`] } },
				],
			});

			// The `simple` format skips a message whose HEADER part is missing.
			const withoutHeader = (uid: number) => ({
				attributes: { uuid: uid, uid, struct: {} },
				parts: [{ which: 'TEXT', body: 'txt' }],
			});

			const runWith = async (staticData: IDataObject, ...batches: unknown[][]) => {
				triggerFunctions.getNode.mockReturnValue(mock<INode>({ typeVersion: 2.1 }));
				triggerFunctions.getNodeParameter.calledWith('format').mockReturnValue('simple');
				triggerFunctions.getNodeParameter.calledWith('downloadAttachments').mockReturnValue(false);
				triggerFunctions.getWorkflowStaticData.mockReturnValue(staticData);

				const connection = connectionReturning(...batches);
				await getNewEmails.call(triggerFunctions, {
					imapConnection: connection,
					searchCriteria: [],
					postProcessAction: 'nothing',
					onEmailBatch: vi.fn(),
				});
				return connection;
			};

			const run = async (staticData: IDataObject, ...batches: unknown[][]) => {
				triggerFunctions.getNode.mockReturnValue(mock<INode>({ typeVersion: 2.1 }));
				triggerFunctions.getNodeParameter.calledWith('format').mockReturnValue('simple');
				triggerFunctions.getNodeParameter.calledWith('downloadAttachments').mockReturnValue(false);
				triggerFunctions.getWorkflowStaticData.mockReturnValue(staticData);

				const emitted: number[] = [];
				const onEmailBatch = vi.fn().mockImplementation((data) => {
					for (const item of data as Array<{ json: { attributes: { uid: number } } }>) {
						emitted.push(item.json.attributes.uid);
					}
				});

				await getNewEmails.call(triggerFunctions, {
					imapConnection: connectionReturning(...batches),
					searchCriteria: [],
					postProcessAction: 'nothing',
					onEmailBatch,
				});

				return emitted;
			};

			it('keeps the watermark below it when it is the last of the batch', async () => {
				const staticData: IDataObject = {};

				const emitted = await run(staticData, [withHeader(10), withoutHeader(11)]);

				expect(emitted).toEqual([10]);
				expect(staticData.lastMessageUid).toBe(10);
			});

			it('keeps the watermark below it even when a later message was emitted', async () => {
				const staticData: IDataObject = {};

				const emitted = await run(staticData, [withHeader(10), withoutHeader(11), withHeader(12)]);

				expect(emitted).toEqual([10, 12]);
				// 12 was emitted, but storing it would hide 11 from every later search.
				expect(staticData.lastMessageUid).toBe(10);
			});

			it('stops at the lowest skipped message when several are skipped', async () => {
				const staticData: IDataObject = {};

				await run(staticData, [
					withHeader(10),
					withoutHeader(11),
					withHeader(12),
					withoutHeader(13),
				]);

				expect(staticData.lastMessageUid).toBe(10);
			});

			it('leaves an existing watermark alone rather than lowering it', async () => {
				const staticData: IDataObject = { lastMessageUid: 50 };

				await run(staticData, [withoutHeader(51)]);

				expect(staticData.lastMessageUid).toBe(50);
			});

			it('still advances the watermark when every message made an item', async () => {
				const staticData: IDataObject = {};

				const emitted = await run(staticData, [withHeader(10), withHeader(11)]);

				expect(emitted).toEqual([10, 11]);
				expect(staticData.lastMessageUid).toBe(11);
			});

			it('still moves the search window past it', async () => {
				// The skipped message is the highest of a full page. The next search
				// must start above it, or the same page comes back for ever.
				const firstPage = [
					...Array.from({ length: 19 }, (_, i) => withHeader(i + 2)),
					withoutHeader(21),
				];

				const connection = await runWith({}, firstPage, []);

				expect(connection.search).toHaveBeenCalledTimes(2);
				expect(connection.search.mock.calls[1][0]).toContainEqual(['UID', '21:*']);
			});

			it('holds the watermark across every page of the same search', async () => {
				const staticData: IDataObject = {};
				// A full page, so the search asks for a second one.
				const firstPage = [
					withoutHeader(1),
					...Array.from({ length: 19 }, (_, i) => withHeader(i + 2)),
				];

				const emitted = await run(staticData, firstPage, [withHeader(21)]);

				// The second page was still fetched and emitted.
				expect(emitted).toContain(21);
				// UID 1 made no item, so no watermark is stored at all and the next
				// search still reaches it.
				expect(staticData.lastMessageUid).toBeUndefined();
			});
		});
	});
});
