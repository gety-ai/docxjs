import {
	Virtualizer,
	elementScroll,
	observeElementOffset,
	observeElementRect
} from '@tanstack/virtual-core';

export interface VirtualizedItem {
	key: string | number;
	estimatedSize: number;
}

interface VirtualizedRendererOptions {
	document: Document;
	hostElement: HTMLElement;
	scrollElement: HTMLElement;
	items: VirtualizedItem[];
	overscan: number;
	renderItem: (index: number) => HTMLElement;
	onRendered?: () => void;
}

export class VirtualizedRenderer {
	private virtualizer: Virtualizer<HTMLElement, HTMLElement>;
	private cleanup: () => void = null;
	private elementCache = new Map<number, HTMLElement>();
	private frameId = 0;
	private topSpacer: HTMLElement;
	private bottomSpacer: HTMLElement;

	constructor(private options: VirtualizedRendererOptions) {
		this.topSpacer = this.createSpacer();
		this.bottomSpacer = this.createSpacer();
		this.virtualizer = new Virtualizer<HTMLElement, HTMLElement>({
			count: options.items.length,
			getScrollElement: () => options.scrollElement,
			estimateSize: index => options.items[index]?.estimatedSize ?? 0,
			scrollToFn: elementScroll,
			observeElementRect,
			observeElementOffset,
			overscan: options.overscan,
			onChange: () => this.scheduleSync()
		});
	}

	mount() {
		this.cleanup = this.virtualizer._didMount();
		this.virtualizer._willUpdate();
		this.sync();
	}

	destroy() {
		if (this.frameId) {
			cancelAnimationFrame(this.frameId);
			this.frameId = 0;
		}

		this.cleanup?.();
		this.cleanup = null;
		this.hostElement.replaceChildren();
		this.elementCache.clear();
	}

	get hostElement() {
		return this.options.hostElement;
	}

	private createSpacer() {
		return Object.assign(this.options.document.createElement("div"), {
			ariaHidden: "true"
		});
	}

	private scheduleSync() {
		if (this.frameId) {
			return;
		}

		this.frameId = requestAnimationFrame(() => {
			this.frameId = 0;
			this.sync();
		});
	}

	private sync() {
		const virtualItems = this.virtualizer.getVirtualItems();
		const nextCache = new Map<number, HTMLElement>();
		const fragment = this.options.document.createDocumentFragment();
		const first = virtualItems[0];
		const last = virtualItems[virtualItems.length - 1];
		const totalSize = this.virtualizer.getTotalSize();

		this.topSpacer.style.height = `${first?.start ?? 0}px`;
		this.topSpacer.style.width = "1px";
		this.bottomSpacer.style.height = `${Math.max(0, totalSize - (last?.end ?? 0))}px`;
		this.bottomSpacer.style.width = "1px";

		fragment.appendChild(this.topSpacer);

		for (const item of virtualItems) {
			const element = this.elementCache.get(item.index) ?? this.options.renderItem(item.index);
			element.dataset.index = `${item.index}`;
			nextCache.set(item.index, element);
			fragment.appendChild(element);
		}

		fragment.appendChild(this.bottomSpacer);
		this.options.hostElement.replaceChildren(fragment);

		for (const item of virtualItems) {
			this.virtualizer.measureElement(nextCache.get(item.index));
		}

		this.elementCache = nextCache;
		this.options.onRendered?.();
	}
}
