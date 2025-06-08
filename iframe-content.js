(function() {
    'use strict';

    // Only run if we're in the PowerApps iframe
    if (!window.location.href.includes('powerapps.com') || window.self === window.top) {
        return;
    }

    console.log("[Extension] Running in PowerApps iframe context");

    let originalParent = null;
    let originalNextSibling = null;
    let originalStyles = {};
    let originalPropertiesStyles = {};
    let isPinned = false;
    let propertiesContainer = null;
    let resizeHandle = null;

    // Polling function to find elements
    function pollForElement(selector, callback, maxAttempts = 30) {
        let attempts = 0;
        console.log("[Extension] Polling for selector: " + selector);
        
        const poll = () => {
            const element = document.querySelector(selector);
            if (element) {
                console.log("[Extension] Found element for selector: " + selector);
                callback(element);
                return;
            }
            
            attempts++;
            if (attempts < maxAttempts) {
                setTimeout(poll, 1000);
            } else {
                console.log("[Extension] Failed to find element after " + maxAttempts + " attempts: " + selector);
            }
        };
        
        poll();
    }

    // Store original styles for reset
    function storeOriginalStyles(element) {
        const computedStyle = window.getComputedStyle(element);
        originalStyles = {
            position: element.style.position || computedStyle.position,
            top: element.style.top || computedStyle.top,
            right: element.style.right || computedStyle.right,
            left: element.style.left || computedStyle.left,
            width: element.style.width || computedStyle.width,
            height: element.style.height || computedStyle.height,
            zIndex: element.style.zIndex || computedStyle.zIndex,
            backgroundColor: element.style.backgroundColor || computedStyle.backgroundColor,
            border: element.style.border || computedStyle.border,
            boxShadow: element.style.boxShadow || computedStyle.boxShadow,
            borderRadius: element.style.borderRadius || computedStyle.borderRadius,
            transform: element.style.transform || computedStyle.transform,
            transition: element.style.transition || computedStyle.transition,
            overflow: element.style.overflow || computedStyle.overflow,
            padding: element.style.padding || computedStyle.padding,
            margin: element.style.margin || computedStyle.margin,
            resize: element.style.resize || computedStyle.resize,
            display: element.style.display || computedStyle.display
        };
        console.log("[Extension] Stored original styles:", originalStyles);
    }

    // Store original properties container styles
    function storeOriginalPropertiesStyles(element) {
        const computedStyle = window.getComputedStyle(element);
        originalPropertiesStyles = {
            transform: element.style.transform || computedStyle.transform,
            transition: element.style.transition || computedStyle.transition,
            marginLeft: element.style.marginLeft || computedStyle.marginLeft,
            left: element.style.left || computedStyle.left,
            right: element.style.right || computedStyle.right,
            zIndex: element.style.zIndex || computedStyle.zIndex
        };
        console.log("[Extension] Stored original properties styles:", originalPropertiesStyles);
    }

    // Enhanced function to find properties container with better logging
    function findPropertiesContainer() {
        console.log("[Extension] Searching for properties container...");
        
        // First, log all sidebar containers to understand the structure
        const allSidebarContainers = document.querySelectorAll('[class*="sidebar-container"]');
        console.log("[Extension] Found", allSidebarContainers.length, "sidebar containers:");
        
        allSidebarContainers.forEach((el, index) => {
            console.log(`[Extension] Sidebar ${index}:`, {
                className: el.className,
                rect: el.getBoundingClientRect(),
                visible: el.offsetWidth > 0 && el.offsetHeight > 0
            });
        });

        // Try various selectors
        const selectors = [
            '.sidebar-container[class*="container_"]',
            '.container_1ma5eibo.sidebar-container',
            '.container_1m5eibo.sidebar-container',
            '.sidebar-container'
        ];
        
        for (const selector of selectors) {
            const containers = document.querySelectorAll(selector);
            console.log(`[Extension] Selector "${selector}" found ${containers.length} elements`);
            
            for (const container of containers) {
                const rect = container.getBoundingClientRect();
                console.log("[Extension] Checking container:", {
                    className: container.className,
                    width: rect.width,
                    height: rect.height,
                    visible: rect.width > 0 && rect.height > 0,
                    right: rect.right
                });
                
                // Return the first visible container that looks like a properties panel
                if (rect.width > 200 && rect.height > 300 && rect.right > window.innerWidth - 50) {
                    console.log("[Extension] Selected this container as properties panel");
                    return container;
                }
            }
        }
        
        console.log("[Extension] No suitable properties container found");
        return null;
    }

    // Get properties container dimensions and position
    function getPropertiesContainerInfo() {
        const container = findPropertiesContainer();
        
        if (container) {
            const rect = container.getBoundingClientRect();
            const computedStyle = window.getComputedStyle(container);
            console.log("[Extension] Properties container selected:", {
                className: container.className,
                width: rect.width,
                height: rect.height,
                top: rect.top,
                right: rect.right
            });
            return {
                element: container,
                width: rect.width,
                height: rect.height,
                top: rect.top,
                right: rect.right,
                computedStyle: computedStyle
            };
        }
        
        return null;
    }

    // Create resize handle
    function createResizeHandle(containerDiv) {
        if (!containerDiv) {
            console.log("[Extension] createResizeHandle: containerDiv is null");
            return null;
        }

        if (resizeHandle) return resizeHandle;

        resizeHandle = document.createElement('div');
        resizeHandle.id = 'formulaBarResizeHandle';
        resizeHandle.style.cssText = `
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 5px;
            background: #e1e1e1;
            cursor: col-resize;
            z-index: 10001;
            border-right: 1px solid #d1d1d1;
        `;

        resizeHandle.addEventListener('mouseenter', () => {
            if (resizeHandle && resizeHandle.style) {
                resizeHandle.style.background = '#0078d4';
            }
        });

        resizeHandle.addEventListener('mouseleave', () => {
            if (resizeHandle && resizeHandle.style) {
                resizeHandle.style.background = '#e1e1e1';
            }
        });

        let isResizing = false;
        let startX = 0;
        let startWidth = 0;

        resizeHandle.addEventListener('mousedown', (e) => {
            isResizing = true;
            startX = e.clientX;
            startWidth = containerDiv.offsetWidth;
            if (document.body.style) {
                document.body.style.cursor = 'col-resize';
                document.body.style.userSelect = 'none';
            }
            e.preventDefault();
        });

        document.addEventListener('mousemove', (e) => {
            if (!isResizing) return;

            const deltaX = startX - e.clientX;
            const newWidth = Math.max(200, Math.min(800, startWidth + deltaX));
            
            if (containerDiv && containerDiv.style) {
                containerDiv.style.setProperty('width', newWidth + 'px', 'important');
                updateFormulaBarSize(containerDiv);
                updatePropertiesPositionInstant(containerDiv); // This should move properties immediately
            }
        });

        document.addEventListener('mouseup', () => {
            if (isResizing) {
                isResizing = false;
                if (document.body.style) {
                    document.body.style.cursor = '';
                    document.body.style.userSelect = '';
                }
            }
        });

        containerDiv.appendChild(resizeHandle);
        return resizeHandle;
    }

    // Update properties position instantly (no transition during resize)
    function updatePropertiesPositionInstant(containerDiv) {
        if (!isPinned) return;
        
        // Re-find properties container in case it changed
        if (!propertiesContainer || !document.body.contains(propertiesContainer)) {
            console.log("[Extension] Properties container not found or removed, searching again...");
            propertiesContainer = findPropertiesContainer();
            if (propertiesContainer) {
                storeOriginalPropertiesStyles(propertiesContainer);
            }
        }
        
        if (!propertiesContainer) {
            console.log("[Extension] No properties container found for instant update");
            return;
        }
        
        const containerRect = containerDiv.getBoundingClientRect();
        const translateX = window.innerWidth - containerRect.left;
        
        console.log("[Extension] Instant update - translateX:", translateX);
        
        if (propertiesContainer.style) {
            // Remove transition for instant movement during resize
            propertiesContainer.style.transition = 'none';
            propertiesContainer.style.transform = `translateX(-${translateX}px)`;
            propertiesContainer.style.zIndex = "999"; // Keep behind formula bar
        }
    }

    // Update properties position with transition (for non-resize operations)
    function updatePropertiesPosition(containerDiv) {
        if (!isPinned) return;
        
        // Re-find properties container in case it changed
        if (!propertiesContainer || !document.body.contains(propertiesContainer)) {
            console.log("[Extension] Properties container not found or removed, searching again...");
            propertiesContainer = findPropertiesContainer();
            if (propertiesContainer) {
                storeOriginalPropertiesStyles(propertiesContainer);
            }
        }
        
        if (!propertiesContainer) {
            console.log("[Extension] No properties container found for position update");
            return;
        }
        
        const containerRect = containerDiv.getBoundingClientRect();
        const translateX = window.innerWidth - containerRect.left;
        
        console.log("[Extension] Updating properties position, translateX:", translateX);
        
        if (propertiesContainer.style) {
            propertiesContainer.style.transition = "transform 0.3s ease-in-out";
            propertiesContainer.style.transform = `translateX(-${translateX}px)`;
            propertiesContainer.style.zIndex = "999"; // Keep behind formula bar
        }
    }

    // Update formula bar size to fit container
    function updateFormulaBarSize(containerDiv) {
        if (!containerDiv) {
            console.log("[Extension] updateFormulaBarSize: containerDiv is null");
            return;
        }

        const containerWidth = containerDiv.offsetWidth;
        const adjustedWidth = Math.max(200, containerWidth - 22); // Account for padding and resize handle

        // Find the formulaBarContainer_w14rc element and override its styles
        const formulaBarContainer = containerDiv.querySelector('.formulaBarContainer_w14rc');
        if (formulaBarContainer) {
            formulaBarContainer.style.setProperty('height', 'auto', 'important');
            formulaBarContainer.style.setProperty('display', 'block', 'important');
            formulaBarContainer.style.setProperty('width', '100%', 'important');
        }

        // Find and override the formulaBarEditor_180503v width
        const formulaBarEditor = containerDiv.querySelector('.formulaBarEditor_180503v, #formulabar.formulaBarEditor_180503v');
        if (formulaBarEditor) {
            formulaBarEditor.style.setProperty('width', adjustedWidth + 'px', 'important');
            formulaBarEditor.style.setProperty('max-width', adjustedWidth + 'px', 'important');
            formulaBarEditor.style.setProperty('min-width', adjustedWidth + 'px', 'important');
            
            // Ensure overflow is visible for intellisense dropdown
            formulaBarEditor.style.setProperty('overflow', 'visible', 'important');
            formulaBarEditor.style.setProperty('position', 'relative', 'important');
            
            console.log("[Extension] Set formulaBarEditor width to " + adjustedWidth + "px");
        }

        // Also target by ID if the class selector doesn't work
        const formulaBarById = containerDiv.querySelector('#formulabar');
        if (formulaBarById) {
            formulaBarById.style.setProperty('width', adjustedWidth + 'px', 'important');
            formulaBarById.style.setProperty('max-width', adjustedWidth + 'px', 'important');
            formulaBarById.style.setProperty('min-width', adjustedWidth + 'px', 'important');
            
            // Ensure overflow is visible for intellisense dropdown
            formulaBarById.style.setProperty('overflow', 'visible', 'important');
            formulaBarById.style.setProperty('position', 'relative', 'important');
            
            console.log("[Extension] Set formulabar by ID width to " + adjustedWidth + "px");
        }

        // Update focusZone to match container
        const focusZone = containerDiv.querySelector('.focusZone-298');
        if (focusZone) {
            focusZone.style.width = '100%';
            focusZone.style.height = '100%';
            focusZone.style.maxWidth = '100%';
            focusZone.style.boxSizing = 'border-box';
        }

        // Fix container overflow to prevent horizontal scrolling
        containerDiv.style.setProperty('overflow-x', 'hidden', 'important');
        containerDiv.style.setProperty('overflow-y', 'auto', 'important');
        
        // Look for Monaco editor container and fix its overflow
        const monacoContainer = containerDiv.querySelector('.monaco-editor');
        if (monacoContainer) {
            monacoContainer.style.setProperty('overflow-x', 'hidden', 'important');
            monacoContainer.style.setProperty('overflow-y', 'visible', 'important');
            
            // Find the monaco editor viewport
            const monacoViewport = monacoContainer.querySelector('.view-lines');
            if (monacoViewport) {
                monacoViewport.style.setProperty('overflow-x', 'hidden', 'important');
                monacoViewport.style.setProperty('overflow-y', 'visible', 'important');
            }
        }

        // Update any formula input areas
        const formulaInputs = containerDiv.querySelectorAll('textarea, input[type="text"], .formula-input, .monaco-editor');
        formulaInputs.forEach(input => {
            if (input && input.style) {
                input.style.width = 'calc(100% - 2px)';
                input.style.maxWidth = 'calc(100% - 2px)';
                input.style.boxSizing = 'border-box';
                
                if (input.tagName.toLowerCase() === 'textarea' || input.tagName.toLowerCase() === 'input') {
                    input.style.setProperty('overflow-x', 'hidden', 'important');
                    input.style.setProperty('overflow-y', 'auto', 'important');
                }
            }
        });

        // Update Monaco editor if present
        const monacoEditor = containerDiv.querySelector('.monaco-editor');
        if (monacoEditor && window.monaco) {
            setTimeout(() => {
                const editor = window.monaco.editor.getEditors().find(e => 
                    e.getDomNode() === monacoEditor
                );
                if (editor) {
                    editor.layout();
                    
                    if (editor.getContribution) {
                        const suggestController = editor.getContribution('editor.contrib.suggestController');
                        if (suggestController && suggestController.widget) {
                            setTimeout(() => {
                                if (suggestController.widget.value) {
                                    suggestController.widget.value.layout();
                                }
                            }, 100);
                        }
                    }
                }
            }, 100);
        }
    }

    // Observe focusZone size changes
    function observeFocusZoneChanges(containerDiv) {
        const focusZone = containerDiv.querySelector('.focusZone-298');
        if (focusZone && window.ResizeObserver) {
            const resizeObserver = new ResizeObserver((entries) => {
                for (let entry of entries) {
                    if (isPinned) {
                        updateFormulaBarSize(containerDiv);
                    }
                }
            });
            
            resizeObserver.observe(focusZone);
            
            // Store observer for cleanup
            containerDiv._focusZoneObserver = resizeObserver;
        }
    }

    // Monitor properties panel visibility and reposition when needed
    function setupPropertiesObserver() {
        if (!isPinned) return;
        
        const observer = new MutationObserver((mutations) => {
            mutations.forEach((mutation) => {
                // Check for new properties containers being added
                mutation.addedNodes.forEach((node) => {
                    if (node.nodeType === Node.ELEMENT_NODE) {
                        // Check if this node or its children contain a sidebar-container
                        const checkForSidebar = (element) => {
                            if (element.classList && element.classList.contains('sidebar-container')) {
                                console.log("[Extension] New sidebar container detected:", element.className);
                                
                                // If this looks like a properties container, reposition it
                                const rect = element.getBoundingClientRect();
                                if (rect.width > 200 && rect.height > 300) {
                                    console.log("[Extension] Properties panel reopened, repositioning immediately");
                                    propertiesContainer = element;
                                    storeOriginalPropertiesStyles(propertiesContainer);
                                    
                                    // Force immediate repositioning without waiting
                                    setTimeout(() => {
                                        repositionPropertiesPanel();
                                    }, 10);
                                    
                                    // Also force another reposition after the panel is fully rendered
                                    setTimeout(() => {
                                        repositionPropertiesPanel();
                                    }, 100);
                                    
                                    return true;
                                }
                            }
                            return false;
                        };
                        
                        // Check the node itself
                        if (checkForSidebar(node)) return;
                        
                        // Check child nodes
                        if (node.querySelectorAll) {
                            const sidebarContainers = node.querySelectorAll('.sidebar-container');
                            sidebarContainers.forEach(container => {
                                if (checkForSidebar(container)) return;
                            });
                        }
                    }
                });
            });
        });

        // Observe the entire document for properties panel changes
        observer.observe(document.body, {
            childList: true,
            subtree: true
        });

        // More aggressive periodic check with immediate repositioning
        const intervalCheck = setInterval(() => {
            if (!isPinned) {
                clearInterval(intervalCheck);
                return;
            }
            
            // Always try to find the current properties container
            const currentPropsContainer = findPropertiesContainer();
            if (currentPropsContainer && currentPropsContainer !== propertiesContainer) {
                console.log("[Extension] Periodic check found new/different properties panel");
                propertiesContainer = currentPropsContainer;
                storeOriginalPropertiesStyles(propertiesContainer);
                
                // Force immediate repositioning
                repositionPropertiesPanel();
                
                // Double-check positioning after a short delay
                setTimeout(() => {
                    repositionPropertiesPanel();
                }, 50);
                
            } else if (currentPropsContainer && propertiesContainer) {
                // Check if positioning was lost
                const currentTransform = currentPropsContainer.style.transform;
                if (!currentTransform || currentTransform === 'none' || currentTransform === '') {
                    console.log("[Extension] Properties panel lost positioning, reapplying");
                    repositionPropertiesPanel();
                }
            }
        }, 100); // Check every 100ms for more responsive updates

        // Store both observer and interval for cleanup
        window._propertiesObserver = observer;
        window._propertiesInterval = intervalCheck;
    }

    // Enhanced reposition function that forces calculation
    function repositionPropertiesPanel() {
        if (!propertiesContainer || !isPinned) {
            console.log("[Extension] repositionPropertiesPanel: propertiesContainer is null or not pinned");
            return;
        }
        
        // Find the current formula bar container to get its position
        const formulaBarContainer = document.querySelector('#formulaBarContainer');
        if (formulaBarContainer) {
            // Force a reflow to ensure we get accurate positioning
            formulaBarContainer.offsetHeight;
            
            const containerRect = formulaBarContainer.getBoundingClientRect();
            const translateX = window.innerWidth - containerRect.left;
            
            console.log("[Extension] Repositioning properties panel, translateX:", translateX, "containerRect:", containerRect);
            
            // Apply the left translation based on formula bar position immediately
            if (propertiesContainer.style) {
                // Apply immediately without transition for instant positioning
                propertiesContainer.style.transition = "none";
                propertiesContainer.style.transform = `translateX(-${translateX}px)`;
                propertiesContainer.style.zIndex = "999"; // Ensure it stays behind formula bar
                
                // Force a style recalculation
                propertiesContainer.offsetHeight;
                
                console.log("[Extension] Applied transform:", propertiesContainer.style.transform);
                
                // Add transition back after positioning
                setTimeout(() => {
                    if (propertiesContainer && propertiesContainer.style) {
                        propertiesContainer.style.transition = "transform 0.3s ease-in-out";
                    }
                }, 10);
            }
        }
    }

    // Apply pinned styles to formula bar
    function pinFormulaBar(containerDiv) {
        console.log("[Extension] Starting pin operation...");
        
        // Find and store properties container
        const propertiesInfo = getPropertiesContainerInfo();
        if (!propertiesInfo) {
            console.log("[Extension] Could not find properties container - aborting pin");
            return;
        }
        
        propertiesContainer = propertiesInfo.element;
        
        // Store original positions
        originalParent = containerDiv.parentNode;
        originalNextSibling = containerDiv.nextSibling;
        storeOriginalStyles(containerDiv);
        storeOriginalPropertiesStyles(propertiesContainer);

        // Get the properties container width for positioning
        const propertiesWidth = propertiesInfo.width;
        const propertiesHeight = propertiesInfo.height;
        
        console.log("[Extension] Moving properties container left by " + propertiesWidth + "px");
        
        // Move properties container to the left with smooth transition
        propertiesContainer.style.transition = "transform 0.3s ease-in-out";
        propertiesContainer.style.transform = "translateX(-" + propertiesWidth + "px)";
        propertiesContainer.style.zIndex = "999"; // Keep behind formula bar

        // Position formula bar in the space where properties used to be
        containerDiv.style.transition = "all 0.3s ease-in-out";
        containerDiv.style.position = "fixed";
        containerDiv.style.top = propertiesInfo.top + "px";
        containerDiv.style.right = "0px";
        containerDiv.style.left = "auto";
        containerDiv.style.setProperty('width', propertiesWidth + 'px', 'important');
        containerDiv.style.setProperty('height', propertiesHeight + 'px', 'important');
        containerDiv.style.setProperty('display', 'block', 'important');
        containerDiv.style.setProperty('overflow', 'visible', 'important');
        containerDiv.style.zIndex = "1000"; // Above properties panel
        containerDiv.style.backgroundColor = "#ffffff";
        containerDiv.style.border = "1px solid #e1e1e1";
        containerDiv.style.boxShadow = "-2px 0 8px rgba(0, 0, 0, 0.1)";
        containerDiv.style.borderRadius = "0";
        containerDiv.style.padding = "16px";
        containerDiv.style.paddingLeft = "21px";
        containerDiv.style.margin = "0";
        containerDiv.style.resize = "none";

        // Move to body to ensure proper positioning
        document.body.appendChild(containerDiv);

        // Create resize handle
        createResizeHandle(containerDiv);

        // Update formula bar sizing and start observing changes
        updateFormulaBarSize(containerDiv);
        observeFocusZoneChanges(containerDiv);

        // Set up mutation observer to catch intellisense dropdowns
        setupIntellisenseObserver();
        
        // Set up observer to monitor properties panel changes
        setupPropertiesObserver();

        isPinned = true;

        // Set up a one-time check to ensure positioning is correct after everything settles
        setTimeout(() => {
            if (isPinned && propertiesContainer) {
                repositionPropertiesPanel();
            }
        }, 1000);

        console.log("[Extension] Formula bar pinned successfully");
    }

    // Set up observer for intellisense/suggest widgets
    function setupIntellisenseObserver() {
        const observer = new MutationObserver((mutations) => {
            mutations.forEach((mutation) => {
                mutation.addedNodes.forEach((node) => {
                    if (node.nodeType === Node.ELEMENT_NODE) {
                        if (node.classList && (
                            node.classList.contains('suggest-widget') ||
                            node.classList.contains('monaco-list') ||
                            node.classList.contains('parameter-hints-widget') ||
                            node.querySelector && node.querySelector('.suggest-widget, .monaco-list, .parameter-hints-widget')
                        )) {
                            console.log("[Extension] Found intellisense widget, fixing positioning");
                            
                            if (node.classList.contains('suggest-widget') || node.classList.contains('monaco-list') || node.classList.contains('parameter-hints-widget')) {
                                fixIntellisenseWidget(node);
                            } else {
                                const widgets = node.querySelectorAll('.suggest-widget, .monaco-list, .parameter-hints-widget');
                                widgets.forEach(fixIntellisenseWidget);
                            }
                        }
                    }
                });
            });
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true
        });

        window._intellisenseObserver = observer;
    }

    // Fix intellisense widget positioning and visibility
    function fixIntellisenseWidget(widget) {
        if (!widget || !widget.style) {
            console.log("[Extension] fixIntellisenseWidget: widget is null or has no style property");
            return;
        }

        widget.style.setProperty('z-index', '10002', 'important');
        widget.style.setProperty('position', 'fixed', 'important');
        widget.style.setProperty('max-height', '300px', 'important');
        widget.style.setProperty('overflow', 'auto', 'important');
        widget.style.setProperty('background', 'white', 'important');
        widget.style.setProperty('border', '1px solid #ccc', 'important');
        widget.style.setProperty('box-shadow', '0 2px 8px rgba(0,0,0,0.15)', 'important');
        
        const textElements = widget.querySelectorAll('.monaco-list-row, .suggest-item, .label-name, .label-text');
        textElements.forEach(el => {
            if (el && el.style) {
                el.style.setProperty('color', '#000', 'important');
                el.style.setProperty('background', 'transparent', 'important');
            }
        });
        
        console.log("[Extension] Applied intellisense widget fixes");
    }

    // Reset formula bar to original position
    function resetFormulaBar(containerDiv) {
        console.log("[Extension] Starting reset operation...");

        // Clean up observers and intervals
        if (window._propertiesObserver) {
            window._propertiesObserver.disconnect();
            delete window._propertiesObserver;
        }

        if (window._propertiesInterval) {
            clearInterval(window._propertiesInterval);
            delete window._propertiesInterval;
        }

        if (window._intellisenseObserver) {
            window._intellisenseObserver.disconnect();
            delete window._intellisenseObserver;
        }

        if (containerDiv._focusZoneObserver) {
            containerDiv._focusZoneObserver.disconnect();
            delete containerDiv._focusZoneObserver;
        }

        // Remove resize handle
        if (resizeHandle) {
            resizeHandle.remove();
            resizeHandle = null;
        }

        // Reset properties container position first
        if (propertiesContainer) {
            console.log("[Extension] Resetting properties container position");
            propertiesContainer.style.transition = "transform 0.3s ease-in-out";
            propertiesContainer.style.transform = originalPropertiesStyles.transform;
            
            // Reset z-index
            if (propertiesContainer.style.zIndex) {
                propertiesContainer.style.removeProperty('z-index');
            }
            
            setTimeout(() => {
                if (propertiesContainer) {
                    Object.keys(originalPropertiesStyles).forEach(prop => {
                        if (originalPropertiesStyles[prop] === '' || originalPropertiesStyles[prop] === 'none') {
                            propertiesContainer.style.removeProperty(prop.replace(/([A-Z])/g, '-$1').toLowerCase());
                        } else {
                            propertiesContainer.style[prop] = originalPropertiesStyles[prop];
                        }
                    });
                }
            }, 300);
        }

        // Reset formula bar styles
        containerDiv.style.transition = "all 0.3s ease-in-out";
        
        setTimeout(() => {
            console.log("[Extension] Restoring formula bar to original parent");
            
            // Remove all custom styles
            Object.keys(originalStyles).forEach(prop => {
                if (originalStyles[prop] === '' || originalStyles[prop] === 'none' || originalStyles[prop] === 'auto') {
                    containerDiv.style.removeProperty(prop.replace(/([A-Z])/g, '-$1').toLowerCase());
                } else {
                    containerDiv.style[prop] = originalStyles[prop];
                }
            });

            // Reset all modified elements
            const elementsToReset = [
                '.focusZone-298',
                '.formulaBarContainer_w14rc', 
                '.formulaBarEditor_180503v',
                '#formulabar'
            ];

            elementsToReset.forEach(selector => {
                const element = containerDiv.querySelector(selector);
                if (element) {
                    ['width', 'height', 'max-width', 'min-width', 'overflow', 'position', 'display', 'box-sizing'].forEach(prop => {
                        element.style.removeProperty(prop);
                    });
                }
            });

            // Move back to original parent
            if (originalParent) {
                if (originalNextSibling && originalNextSibling.parentNode === originalParent) {
                    originalParent.insertBefore(containerDiv, originalNextSibling);
                } else {
                    originalParent.appendChild(containerDiv);
                }
            }

            isPinned = false;
            propertiesContainer = null;

            console.log("[Extension] Reset operation completed");
        }, 100);
    }

    // Inject pin and reset buttons
    function injectButtons(containerDiv) {
        if (!containerDiv.querySelector("#pinFormulaBarButton")) {
            console.log("[Extension] Injecting Pin to Right button into container.");
            
            const pinIconSvg = `
                <svg viewBox='0 0 2048 2048' xmlns='http://www.w3.org/2000/svg' style="width: 16px; height: 16px; fill: currentColor;">
                    <path d='M1963 512h85v853h-85q-103 0-190-44t-150-126h-317q-37 78-93 141t-127 107-151 69-167 24h-85v-512H171L0 939l171-86h512V341h85q86 0 167 24t151 69 127 108 93 141h317q62-81 149-126t191-45z'></path>
                </svg>
            `;
            
            const unpinIconSvg = `
                <svg viewBox='0 0 2048 2048' xmlns='http://www.w3.org/2000/svg' style="width: 16px; height: 16px; fill: currentColor;">
                    <path d='M2048 512v853h-85q-38 0-75-8t-73-21L825 345q78 8 150 35t134 71 113 103 84 129h317q60-81 150-126t190-45h85zM25 146L146 25l1877 1877-121 121-689-689q-42 48-93 85t-108 64-118 39-126 14h-85v-512H171L0 939l171-86h512v-50L25 146z'></path>
                </svg>
            `;
            
            // Match the PowerApps button styling
            const buttonHTML = `
                <button id="pinFormulaBarButton" class="button_18dzwa2-o_O-focus_x46kcf" data-is-focusable="true" style="
                    margin-left: 8px;
                    padding: 4px 8px;
                    cursor: pointer;
                    background: transparent;
                    color: #323130;
                    border: 1px solid transparent;
                    border-radius: 2px;
                    font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Roboto', 'Helvetica Neue', sans-serif;
                    font-size: 14px;
                    font-weight: 400;
                    min-height: 32px;
                    height: 32px;
                    min-width: 32px;
                    outline: transparent;
                    position: relative;
                    text-align: center;
                    text-decoration: none;
                    user-select: none;
                    vertical-align: top;
                    display: inline-flex;
                    align-items: center;
                    justify-content: center;
                    box-sizing: border-box;
                    transition: all 0.1s ease 0s;
                " title="Pin formula bar to right side" aria-label="Pin formula bar to right side">
                    ${pinIconSvg}
                </button>
            `;
            
            containerDiv.insertAdjacentHTML('beforeend', buttonHTML);

            const pinButton = containerDiv.querySelector("#pinFormulaBarButton");
            
            // Add PowerApps-style hover and focus effects
            pinButton.addEventListener('mouseenter', () => {
                if (!isPinned) {
                    pinButton.style.backgroundColor = '#f3f2f1';
                    pinButton.style.borderColor = '#8a8886';
                } else {
                    pinButton.style.backgroundColor = '#edebe9';
                    pinButton.style.borderColor = '#8a8886';
                }
            });
            
            pinButton.addEventListener('mouseleave', () => {
                if (!isPinned) {
                    pinButton.style.backgroundColor = 'transparent';
                    pinButton.style.borderColor = 'transparent';
                } else {
                    pinButton.style.backgroundColor = '#f3f2f1';
                    pinButton.style.borderColor = '#8a8886';
                }
            });

            pinButton.addEventListener('mousedown', () => {
                pinButton.style.backgroundColor = '#edebe9';
                pinButton.style.borderColor = '#8a8886';
            });

            pinButton.addEventListener('mouseup', () => {
                if (!isPinned) {
                    pinButton.style.backgroundColor = '#f3f2f1';
                } else {
                    pinButton.style.backgroundColor = '#f3f2f1';
                }
            });

            pinButton.addEventListener("click", (e) => {
                e.preventDefault();
                e.stopPropagation();
                
                if (!isPinned) {
                    console.log("[Extension] Pin to Right button clicked.");
                    pinFormulaBar(containerDiv);
                    // Update button appearance for pinned state
                    pinButton.innerHTML = unpinIconSvg;
                    pinButton.style.backgroundColor = "#f3f2f1";
                    pinButton.style.borderColor = "#8a8886";
                    pinButton.title = "Reset formula bar position";
                    pinButton.setAttribute('aria-label', 'Reset formula bar position');
                } else {
                    console.log("[Extension] Reset Position button clicked.");
                    resetFormulaBar(containerDiv);
                    // Update button appearance for unpinned state
                    pinButton.innerHTML = pinIconSvg;
                    pinButton.style.backgroundColor = "transparent";
                    pinButton.style.borderColor = "transparent";
                    pinButton.title = "Pin formula bar to right side";
                    pinButton.setAttribute('aria-label', 'Pin formula bar to right side');
                }
            });
        }
    }

    // Initialize the extension within iframe
    function initializeInIframe() {
        console.log("[Extension] Initializing formula bar pinner within PowerApps iframe.");
        
        // Look for the formula bar container
        pollForElement("#formulaBarContainer", (containerDiv) => {
            console.log("[Extension] Found formula bar container:", containerDiv);
            injectButtons(containerDiv);
        });
    }

    // Start initialization
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initializeInIframe);
    } else {
        setTimeout(initializeInIframe, 1000);
    }

})();