from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Create document
doc = Document()

# Set margins
sections = doc.sections
for section in sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

# Add title
title = doc.add_heading('SportZone Frontend', 0)
title_run = title.runs[0]
title_run.font.color.rgb = RGBColor(255, 102, 0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

subtitle = doc.add_heading('Complete Interview Questions & Answers (100 Q&A)', level=2)
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Add metadata
date_para = doc.add_paragraph('Generated: January 31, 2026')
date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
date_para.runs[0].font.size = Pt(10)
date_para.runs[0].font.color.rgb = RGBColor(128, 128, 128)

doc.add_paragraph()

# All 100 Q&A pairs
questions_answers = [
    ("What is the overall architecture of the SportZone frontend application?", "Single Page Application (SPA) built with React 18.2.0 using Vite. Component-based architecture with Context API for theme, React Router v6 for routing, Axios with interceptors for API calls, Bootstrap 5.3.2 for styling, and JWT token-based authentication with localStorage persistence. Role-based access control (USER, VENUE_OWNER, ADMIN)."),
    ("Explain the component hierarchy and data flow.", "App.jsx root with ThemeProvider wrapping Navigation and Routes. Data flows: user interaction → state update → API call → response → state update → re-render. Props flow downward, events upward via callbacks. Global state (theme) uses Context API."),
    ("What are the main pages and their responsibilities?", "Home: Landing page. VenueList: Browse venues with filters. VenueDetails: Booking workflow. BookingForm: Time slot selection. UserBookings: History with pagination. UserProfile: Update name/password. AdminDashboard: System admin. VenueOwnerDashboard: Manage venues. VenueDashboard: Individual venue management. Login/Signup: Authentication."),
    ("How does authentication work in SportZone?", "User enters credentials → POST to /api/users/login → Backend returns JWT + user object → Frontend stores in localStorage → Axios interceptor adds Authorization header to all requests → Protected routes check localStorage user presence and role."),
    ("What role-based access control (RBAC) is implemented?", "Three roles: USER (browse, book), VENUE_OWNER (manage venues), ADMIN (manage all). ProtectedRoute component checks user.role, redirects unauthorized. Backend must enforce RBAC on API endpoints."),
    ("Explain the BookingForm component in detail.", "Manages time slot selection with smart logic: empty → startTime selected → toggle or extend → calculate price. State: selectedSport, selectedDate, startTime, endTime, selectedCourt, bookedSlots. Validates: future date, no conflicts, hourly increments. Fetches booked slots per court/date."),
    ("How does payment integration with Razorpay work?", "Step 1: Create booking (PENDING). Step 2: Create Razorpay order. Step 3: Open payment modal, user pays, verify signature on backend. If valid → confirm booking, if invalid → delete booking. Signature verification prevents tampering."),
    ("Explain the theme management system.", "ThemeContext with useState('light'/'dark'). useEffect sets document data-theme attribute. CSS variables in :root and [data-theme]. Instant switching via CSS cascade. Persists to localStorage."),
    ("How does Navigation handle different user roles?", "Base: logo, Home, Venues. Conditional: Login/Signup (not logged in), Profile/MyBookings/Logout (logged in), Dashboard (ADMIN), Owner Dashboard (VENUE_OWNER). Always: theme toggle. Checks localStorage user.role."),
    ("Explain the ProtectedRoute component.", "Checks localStorage for user. If missing → redirect to /login. If role provided and mismatch → redirect to /. Otherwise render children. Frontend protection only; backend must enforce."),
    ("What is axiosConfig.js and how does it work?", "Request interceptor that reads user from localStorage, extracts token, adds Authorization header: 'Bearer {token}' to all requests. No manual header addition needed. Limitations: XSS vulnerability, no token expiration, no refresh mechanism."),
    ("What are the main security vulnerabilities?", "localStorage XSS vulnerability, no token expiration, frontend validation insufficient, price tampering possible, double-booking race conditions, payment failure handling incomplete."),
    ("How would you secure sensitive data storage?", "Use httpOnly cookies instead of localStorage, implement token expiration with refresh tokens, add CSRF protection, validate all input on backend, sanitize output, use HTTPS only."),
    ("Explain JWT token structure and verification.", "JWT: header.payload.signature. Header: algorithm (HS256). Payload: user claims (id, role, exp). Signature: HMAC(header+payload, secret). Backend verifies signature matches, checks expiration, validates claims."),
    ("How is XSS prevention implemented?", "React auto-escapes JSX output by default. No dangerouslySetInnerHTML used. Input sanitization important. CSP headers on backend. Avoid eval and dynamic code execution."),
    ("What CORS configuration is needed?", "Backend allows requests from frontend domain. Headers: Access-Control-Allow-Origin, Allow-Methods, Allow-Headers, Allow-Credentials. Credentials: true if sending cookies. Preflight OPTIONS request for complex requests."),
    ("Explain API versioning strategy.", "Path-based: /api/v1/venues vs /api/v2/venues. Support 2-3 versions simultaneously. Deprecation timeline: announce → show warning headers → sunset. Breaking changes require version bump. Backward compatible changes (adding fields) don't."),
    ("How would you handle API rate limiting?", "Token bucket algorithm for client-side limits. Request queuing for concurrency control. Exponential backoff retry. Debounce/throttle user inputs. Rate-Limit headers from API. Show user messaging when limited."),
    ("What testing approaches are recommended?", "Unit tests: individual functions, hooks. Integration tests: components with mocked API. E2E tests: full user workflows. Mocking: jest.mock() for axios, localStorage. Coverage target: 80%+."),
    ("How would you debug API issues?", "Network tab in DevTools. Log request/response. Check headers. Verify JSON format. Backend logs. Postman for testing. Mock API responses. Check CORS. Verify authentication token."),
    ("Explain state management with Context API.", "Create context with createContext(). Provider component with state and value prop. useContext(ThemeContext) in components. Simple for global state. Prop drilling eliminated. Performance: all subscribers re-render on any state change."),
    ("What are Context API limitations?", "All subscribers re-render (performance issue for large trees). No built-in selectors. No middleware support. No time-travel debugging. Better for small, infrequently changing state."),
    ("How would you optimize Component re-renders?", "React.memo() for functional components. useMemo() for expensive computations. useCallback() for stable function references. Separate contexts for different state types. Virtual scrolling for large lists."),
    ("Explain lazy loading and code splitting.", "React.lazy() for dynamic imports. Suspense boundary with fallback. Routes: const AdminDashboard = lazy(() => import('./AdminDashboard')). Reduces initial bundle. Async chunk loading."),
    ("How would you handle form validation?", "Client-side: HTML5 validation, regex patterns, length checks. Show errors inline. Server-side: ALL validation must happen (not just frontend). Prevent type tampering, SQL injection, XSS."),
    ("Explain carousel implementation.", "setInterval with state for current index. Increment index, wrap at end. Controls: prev/next buttons, dot indicators. Fade or slide transition. Auto-rotation every 2-5 seconds. Image onError fallback."),
    ("What pagination patterns are used?", "API: page, size parameters. Frontend: track current page, total. Previous/Next buttons. Offset: page * size. Limit: page size. Cursor-based pagination better for large datasets."),
    ("How is modal/dialog handled?", "State: isOpen boolean. Conditional render. Overlay for backdrop. Close on X button or outside click. Prevent body scroll when modal open (overflow: hidden). Focus trap for accessibility."),
    ("Explain image optimization.", "Compression: reduce file size 70-80%. Formats: WebP > JPEG > PNG. Responsive: multiple sizes for screen widths. Lazy loading: load when visible. CDN: serve from edge servers. Progressive JPEG: low-quality placeholder."),
    ("What is sticky positioning used for?", "BookingForm sticky on VenueDetails page. User sees booking form while scrolling venue info. CSS: position: sticky; top: 20px. Doesn't remove from document flow. Respects scroll container boundaries."),
    ("How would you implement real-time notifications?", "WebSocket for two-way communication. Server publishes events (booking.created). Client subscribes, receives instantly. Alternative: SSE (one-way), Polling (wasteful), FCM (push notifications)."),
    ("Explain client-side caching strategy.", "Memory cache: component state, short TTL. localStorage: 5MB limit, persists. Service Worker: network-first, cache-first strategies. SWR pattern: serve stale immediately, fetch fresh. Manual invalidation on mutations."),
    ("What is stale-while-revalidate pattern?", "Return cached data immediately. Fetch fresh in background. Update cache when arrives. Perceived speed improvement. Network efficient. Works offline. Best for infrequently changing data."),
    ("How would you handle offline scenarios?", "Service Worker caches responses. Sync queue: store actions while offline. localStorage for data persistence. IndexedDB for complex data. Background sync when online. Show offline indicator UI."),
    ("Explain eventual consistency concept.", "Data eventually consistent but temporarily inconsistent. Optimistic updates: assume success, revert if failed. Version numbers for conflict detection. Merge strategies. User communication important."),
    ("What is the error boundary component?", "Catches errors from child components. Prevents app crash. Shows fallback UI. Logs error. Doesn't catch event handlers, async code. Different boundaries at different levels for granular error handling."),
    ("How would you implement feature flags?", "Config object with feature toggles. Backend-driven flags: API returns enabled features. Gradual rollout: percentage of users. A/B testing: random assignment per user. Admin panel to toggle flags."),
    ("Explain A/B testing implementation.", "Random user assignment to variant A or B. Hash userId for consistency. Metric tracking: conversions, time, etc. Statistical significance testing. Gradual rollout: 5% → 25% → 50% → 100%."),
    ("What production monitoring is needed?", "Error tracking (Sentry). Logging service (LogRocket). Performance monitoring (Lighthouse CI). User analytics. Uptime monitoring. Alerts for errors, degraded performance, payment failures."),
    ("How would you handle long-running operations?", "setTimeout for UI breathing room. Web Workers for CPU-intensive tasks. Backend async jobs with polling. Streaming/chunking for large uploads. Show progress indicators."),
    ("Explain security best practices.", "HTTPS only. CSP headers. Input validation/sanitization. Output encoding. XSS prevention. CSRF tokens. httpOnly cookies. Rate limiting. SQL injection prevention. Principle of least privilege."),
    ("What accessibility features should be implemented?", "Semantic HTML. ARIA labels. Keyboard navigation. Color contrast (WCAG AA). Alt text for images. Focus indicators. Form labels. Screen reader compatibility."),
    ("How would you implement undo/redo?", "Command history stack. Array of states. Index for current position. Undo: pop from history. Redo: push from future. Memory management: limit history size. UI controls to show availability."),
    ("Explain collaborative features (real-time editing).", "CRDT or Operational Transform for consistency. WebSocket for real-time sync. Cursor presence (see other users). Conflict resolution: merge strategies. Offline support with sync queue."),
    ("What is progressive web app (PWA)?", "Service Worker for offline. Web manifest for installation. Push notifications. Background sync. App-like experience. Installable on home screen. Works offline or slow networks."),
    ("How would you implement two-factor authentication?", "TOTP with Google Authenticator. QR code generation. Backup codes for recovery. Email/SMS verification alternative. Secure secret storage (hashed). Rate limiting on verification attempts."),
    ("Explain microservices architecture.", "Split monolith into independent services. API Gateway routes requests. Message Queue for async communication. Each service owns database. Benefits: scalability, independence. Challenges: debugging, consistency."),
    ("What is database indexing impact?", "Index accelerates queries 10-100x. B-tree structure for fast lookup. Trade-off: faster reads, slower writes. Index maintenance overhead. Choose columns in WHERE, JOIN, ORDER BY. Monitor unused indexes."),
    ("How would you handle white-label customization?", "Dynamic branding via config. CSS variables for colors, logos. Conditional components per brand. Multi-tenant database schema. Separate domains per brand. Admin panel for customization."),
    ("Explain API versioning timeline.", "Release v2 alongside v1. Support both for 6-12 months. Document migration path. Send notifications to v1 users. Deprecation warnings in headers. Hard sunset date, return 404."),
    ("What deployment strategies exist?", "Blue-green: two environments, switch traffic. Canary: gradual rollout to users. Rolling: sequential instance updates. Shadow: duplicate traffic to new version. Feature flags for gradual rollout."),
    ("How would you optimize React performance?", "Memoization: React.memo(), useMemo(), useCallback(). Code splitting: lazy loading. Virtual scrolling. Profiling with DevTools. Remove unused dependencies. Optimize bundle size."),
    ("Explain observability (logs, metrics, traces).", "Logs: what happened (error messages, events). Metrics: quantitative (CPU, requests/sec, errors/min). Traces: request path through services. Together provide complete visibility. Tools: ELK, Prometheus, Jaeger."),
    ("What field-level permissions mean?", "User role determines visible fields. Admin sees email, internal notes. Customer sees only own info. Frontend hiding isn't secure. Backend must enforce on API responses."),
    ("How would you prepare for production launch?", "Load testing at 10x traffic. Security audit. Disaster recovery plan. Monitoring setup. On-call support. Documentation. User education. Gradual rollout: internal → beta → public."),
    ("What are common performance bottlenecks?", "Large bundle size. Unoptimized images. N+1 queries. Synchronous API calls. Unmemoized components. Missing indexes. Inefficient database queries. Network waterfall."),
    ("Explain bundle size optimization.", "Tree-shaking: remove unused code. Code splitting: lazy routes. Minification: compress code. Compression: gzip, brotli. Analyze: webpack-bundle-analyzer. Defer non-critical JS."),
    ("How would you implement service mesh?", "Istio, Linkerd, Consul. Sidecar proxies intercept traffic. Traffic management, security, monitoring. Circuit breaking, retry logic. Service discovery. Setup complexity overhead."),
    ("Explain chaos engineering approach.", "Intentionally break things in production. Kill services, introduce latency, corrupt data. Observe resilience. Fix weaknesses. Learn failure modes. Prevents surprise outages."),
    ("What is header-based CORS?", "Version in Accept header: Accept: application/vnd.sportzone.v2+json. Server checks header, routes. Cleaner URLs. Requires documentation."),
    ("How would you handle payment failures?", "Retry logic with exponential backoff. Store failed transactions. Alert user, show retry option. Idempotency key prevents double-charging. Reconciliation: verify payment status."),
    ("Explain session management best practices.", "Short expiration (15-30 min). Refresh tokens for renewal. Secure cookies (httpOnly, Secure, SameSite). CSRF tokens. Session invalidation on logout. Backend validates all requests."),
    ("What is contract testing in APIs?", "Consumer tests API contracts. Provider tests implementation. Prevents integration breakage. Both sides verify expectations. PACT testing framework."),
    ("How would you implement feature completeness tracking?", "Dashboard: % complete per feature. Checklist: API endpoints, screens, tests, docs. Status: red/yellow/green. Release when all green. Burndown chart."),
    ("Explain circuit breaker pattern.", "Service calls external service. Monitor failures. Open circuit: fail fast (don't call). Half-open: test if recovered. Closed: normal operation. Prevents cascading failures."),
    ("What is dependency injection?", "Inject dependencies rather than creating. Constructor injection or provider. Loosely coupled code. Easier testing with mocks. Frameworks: Spring, Angular (built-in)."),
    ("How would you handle vendor lock-in?", "Abstract payment provider behind interface. Multiple implementations available. Easy swap. Cloud provider abstraction layer. Open standards when possible."),
    ("Explain canary deployment.", "Deploy to small % of servers (5%). Monitor metrics. Gradually increase (25%, 50%, 100%). Automatic rollback if errors spike. Low-risk rollout."),
    ("What is shadow traffic?", "Duplicate production traffic to new version. New version doesn't affect users. Compare metrics: latency, errors, behavior. Validate before full rollout."),
    ("How would you implement distributed tracing?", "Unique ID per request. Pass through all services. Each service logs with trace ID. Collect logs centrally. Visualize request path. Tools: Jaeger, Zipkin."),
    ("Explain circuit breaker states.", "Closed: normal (count failures). Open: fail fast (blocked, timeout). Half-open: test (allow one request). Success → Closed. Failure → Open again."),
    ("What is API contract versioning?", "Document expected request/response format. Version contracts. Notify when breaking change. Consumer-driven contract testing. Both sides must agree on contract."),
    ("How would you prevent N+1 query problem?", "Batch loading: load all related items once. Join queries in DB. GraphQL DataLoader. Eager loading. Pagination to limit result sets. Index foreign keys."),
    ("Explain webhook implementation.", "App A sends HTTP POST to App B when event occurs. Retry logic for failures. Signature verification for security. Idempotency key for duplicate prevention. Status page for subscription management."),
    ("What is event sourcing?", "Store all events (bookings, payments). Current state derived from events. Event replay for debugging. Temporal queries: state at any point. Event versioning for evolution."),
    ("How would you implement role hierarchies?", "Admin > Venue Owner > User. Role inheritance (inheritance chain). Permission matrix (role + action). Check highest role permission. Easier than flat roles."),
    ("Explain exponential backoff retry.", "First retry: 1 sec. Second: 2 sec. Third: 4 sec. Pattern: delay = base * (2^attempt). Max jitter to prevent thundering herd. Max retries limit."),
    ("What is rate-limit bucket strategy?", "Tokens: initial amount, refill per second. Request costs token(s). Enough tokens → proceed. Else → queue/reject. Smooth traffic handling."),
    ("How would you handle concurrent editing?", "Locks: pessimistic (block), optimistic (version check). CRDT for conflict-free. Merge strategies. User notification on conflicts. Undo/redo for changes."),
    ("Explain semantic versioning.", "MAJOR.MINOR.PATCH (v1.2.3). MAJOR: breaking changes. MINOR: backward-compatible features. PATCH: bug fixes. Pre-release: v1.0.0-alpha. Build metadata: v1.0.0+build.123."),
    ("What is strangler fig pattern?", "Gradually replace monolith with microservices. API Gateway intercepts calls. Route old paths to monolith, new to services. Over time, migrate more. Minimize disruption."),
    ("How would you implement consistent hashing?", "Map keys to ring. Nodes placed on ring. Key → nearest node clockwise. Node failure: keys reassigned to next. Minimal rehashing. Load balancing: multiple nodes."),
    ("Explain bloom filter use cases.", "Check if item in set (fast, probabilistic). False positives possible, no false negatives. Memory efficient. Example: cache bloom filter of seen requests (prevent replays)."),
    ("What is token refresh mechanism?", "Access token short-lived (15 min). Refresh token long-lived (7 days). Access token expired → request new with refresh token. Backend validates refresh token → issues new access token. Refresh token also expires eventually."),
    ("How would you handle decimal precision?", "Prices as integers: store 500 (= $5.00). Avoid floating-point for money. Rounding errors prevented. Work with cents/paise, convert for display."),
    ("Explain gossip protocol.", "Nodes share state with random peers. Information spreads exponentially. Handles network partitions. Eventually consistent. Used in: Cassandra, Consul."),
    ("What is bulkhead pattern?", "Isolate failures. Separate thread pools for different services. One service fails → doesn't affect others. Threads exhausted → others continue. Prevents cascading failure."),
    ("How would you implement search autocomplete?", "Trie data structure for fast prefix lookup. Typeahead API: GET /search?q=cricket&limit=10. Cache frequent searches. Limit results (top 10). Fuzzy matching for typos."),
    ("Explain GDPR compliance requirements.", "Right to access: user can download data. Right to be forgotten: delete all data. Consent: explicit opt-in. Data minimization: collect only needed. Privacy policy required."),
    ("What is database sharding?", "Partition data across multiple servers. Shard key determines server. User bookings on shard by userId. Increases capacity. Requires application awareness. Query routing complexity."),
    ("How would you prevent timing attacks?", "Constant-time comparison: check entire input, not short-circuit. String length comparison timing attacks. Use bcrypt, argon2 (slow by design). Avoid: if (password == inputPassword)."),
    ("Explain read replicas use case.", "Primary DB: writes. Replica DBs: reads only. Distribute read load. Asynchronous replication (eventual consistency). Failover to replica if primary fails. Query routing: write to primary, read to replicas."),
    ("What is eventual consistency guarantee?", "Reads might return stale data temporarily. Eventually converge. Acceptable for analytics, recommendations. Unacceptable for: financial transfers, inventory (needs strong consistency)."),
    ("How would you implement geo-redundancy?", "Multiple data centers in different regions. Replicate data globally. Route users to nearest DC. Handles regional outages. Higher latency for replication. Cost overhead."),
]

# Build document with all Q&A
for i, (question, answer) in enumerate(questions_answers, 1):
    # Question
    q_heading = doc.add_heading(f'Q{i}. {question}', level=2)
    q_heading.runs[0].font.color.rgb = RGBColor(0, 102, 204)
    
    # Answer
    answer_para = doc.add_paragraph(answer)
    answer_para.paragraph_format.space_after = Pt(12)
    
    # Page break every 10 questions for readability
    if i % 10 == 0 and i < len(questions_answers):
        doc.add_page_break()

# Save
output_path = r'd:\SportZone\SportZone_Interview_Questions_Complete_100.docx'
doc.save(output_path)

print("✓ Word document created successfully!")
print(f"✓ Location: {output_path}")
print(f"✓ Total Questions: {len(questions_answers)}")
print(f"✓ File is ready for download!")
