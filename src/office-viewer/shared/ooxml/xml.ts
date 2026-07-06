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

  // 首先尝试标准方法
  const standardMatch = Array.from(node.getElementsByTagName('*')).find(
    (child) => normalizedLocalName(child) === normalized || child.localName.toLowerCase() === normalized,
  );

  if (standardMatch) {
    return standardMatch;
  }

  // 对于 VML 元素（如 v:textbox），尝试使用命名空间 URI 查找
  // VML 命名空间: urn:schemas-microsoft-com:vml
  // WordprocessingML 命名空间: http://schemas.microsoft.com/office/word/2003/wordml
  const vmlNamespaces = [
    'urn:schemas-microsoft-com:vml',
    'http://schemas.microsoft.com/office/word/2003/wordml',
  ];

  for (const ns of vmlNamespaces) {
    try {
      const nsMatch = node.getElementsByTagNameNS(ns, normalized);
      if (nsMatch && nsMatch.length > 0) {
        return nsMatch[0];
      }
    } catch {
      // 某些环境可能不支持 getElementsByTagNameNS，忽略错误
    }
  }

  // 最后尝试通过完整的标签名查找（包括常见的 VML 前缀）
  const prefixes = ['v:', 'w:', 'o:'];
  for (const prefix of prefixes) {
    try {
      const prefixedMatch = node.getElementsByTagName(prefix + normalized);
      if (prefixedMatch && prefixedMatch.length > 0) {
        return prefixedMatch[0];
      }
    } catch {
      // 忽略错误
    }
  }

  return null;
}

export function descendantsByLocalName(node: Element | null | undefined, localName: string) {
  if (!node) {
    return [];
  }

  const normalized = localName.toLowerCase();

  // 首先尝试标准方法
  const standardMatches = Array.from(node.getElementsByTagName('*')).filter(
    (child) => normalizedLocalName(child) === normalized || child.localName.toLowerCase() === normalized,
  );

  if (standardMatches.length > 0) {
    return standardMatches;
  }

  // 对于 VML 元素，尝试使用命名空间 URI 查找
  const vmlNamespaces = [
    'urn:schemas-microsoft-com:vml',
    'http://schemas.microsoft.com/office/word/2003/wordml',
  ];

  for (const ns of vmlNamespaces) {
    try {
      const nsMatches = Array.from(node.getElementsByTagNameNS(ns, normalized));
      if (nsMatches.length > 0) {
        return nsMatches;
      }
    } catch {
      // 某些环境可能不支持 getElementsByTagNameNS，忽略错误
    }
  }

  // 最后尝试通过完整的标签名查找（包括常见的 VML 前缀）
  const prefixes = ['v:', 'w:', 'o:'];
  for (const prefix of prefixes) {
    try {
      const prefixedMatches = Array.from(node.getElementsByTagName(prefix + normalized));
      if (prefixedMatches.length > 0) {
        return prefixedMatches;
      }
    } catch {
      // 忽略错误
    }
  }

  return [];
}
