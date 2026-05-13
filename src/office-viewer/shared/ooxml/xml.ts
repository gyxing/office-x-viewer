export function parseXml(xml: string) {
  if (typeof DOMParser === 'undefined') {
    throw new Error('DOMParser is not available in this environment');
  }
  try {
    return new DOMParser().parseFromString(xml, 'application/xml');
  } catch {
    const document = new DOMParser().parseFromString(xml, 'text/html');
    const root =
      Array.from(document.body?.children ?? []).find((node) => node.nodeType === 1) ??
      Array.from(document.children ?? []).find((node) => node.nodeType === 1) ??
      document.documentElement;
    return new Proxy(document, {
      get(target, prop, receiver) {
        if (prop === 'documentElement') return root;
        if (prop === 'querySelector') return root?.querySelector?.bind(root) ?? target.querySelector.bind(target);
        if (prop === 'querySelectorAll') return root?.querySelectorAll?.bind(root) ?? target.querySelectorAll.bind(target);
        if (prop === 'getElementsByTagName') {
          return root?.getElementsByTagName?.bind(root) ?? target.getElementsByTagName.bind(target);
        }
        return Reflect.get(target, prop, receiver);
      },
    });
  }
}

export function textContent(node: Element | null | undefined) {
  return node?.textContent ?? '';
}

export function attr(node: Element | null | undefined, name: string) {
  if (!node) {
    return undefined;
  }

  const direct = node.getAttribute(name);
  if (direct !== null) {
    return direct;
  }

  const localName = name.includes(':') ? name.split(':').pop() : name;
  const attributes = node.attributes ? Array.from(node.attributes) : [];
  const matched = attributes.find((item) => item.localName === localName || item.name === name);
  return matched?.value;
}

function normalizedLocalName(node: Element) {
  return node.localName.split(':').pop()?.toLowerCase() ?? node.localName.toLowerCase();
}

export function matchesLocalName(node: Element | null | undefined, localName: string) {
  if (!node) return false;
  return normalizedLocalName(node) === localName.toLowerCase();
}

export function descendantByXmlLocalName(node: Element | null | undefined, localName: string) {
  if (!node) return null;
  const normalized = localName.toLowerCase();
  return (
    Array.from(node.getElementsByTagName('*')).find(
      (child) => child.nodeName.includes(':') && normalizedLocalName(child) === normalized,
    ) ?? null
  );
}

export function firstElement(node: Element | null | undefined, selector: string) {
  return node?.querySelector(selector) ?? null;
}

export function allElements(node: Element | null | undefined, selector: string) {
  return node ? Array.from(node.querySelectorAll(selector)) : [];
}

export function childByLocalName(node: Element | null | undefined, localName: string) {
  if (!node) {
    return null;
  }

  const normalized = localName.toLowerCase();
  return Array.from(node.children).find((child) => normalizedLocalName(child) === normalized || child.localName.toLowerCase() === normalized) ?? null;
}

export function childrenByLocalName(node: Element | null | undefined, localName: string) {
  if (!node) {
    return [];
  }

  const normalized = localName.toLowerCase();
  return Array.from(node.children).filter((child) => normalizedLocalName(child) === normalized || child.localName.toLowerCase() === normalized);
}

export function descendantByLocalName(node: Element | null | undefined, localName: string) {
  if (!node) {
    return null;
  }

  const normalized = localName.toLowerCase();
  return Array.from(node.getElementsByTagName('*')).find(
    (child) => normalizedLocalName(child) === normalized || child.localName.toLowerCase() === normalized,
  ) ?? null;
}

export function descendantsByLocalName(node: Element | null | undefined, localName: string) {
  if (!node) {
    return [];
  }

  const normalized = localName.toLowerCase();
  return Array.from(node.getElementsByTagName('*')).filter(
    (child) => normalizedLocalName(child) === normalized || child.localName.toLowerCase() === normalized,
  );
}
