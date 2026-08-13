
    function getPDFOptions(filename) {
      return {
        margin:       [10, 10, 10, 10],
        filename:     filename,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  {
          scale: 2,
          useCORS: true,
          logging: false,
          scrollX: 0,
          scrollY: 0,
          width: 960,
          windowWidth: 960,
          x: 0,
          y: 0,
          onclone: function(clonedDoc) {
            if (clonedDoc.defaultView && clonedDoc.defaultView.frameElement) {
              clonedDoc.defaultView.frameElement.style.width = "960px";
              clonedDoc.defaultView.frameElement.style.minWidth = "960px";
            }
            clonedDoc.documentElement.style.cssText = "width:960px!important;min-width:960px!important;overflow:visible!important;";
            clonedDoc.body.style.cssText = "width:960px!important;min-width:960px!important;margin:0!important;padding:0!important;overflow:visible!important;background:#ffffff!important;";
            var clonedEl = clonedDoc.getElementById("pdf-report-content");
            if (clonedEl) {
              clonedEl.style.cssText = "width:960px!important;max-width:960px!important;min-width:960px!important;background:#ffffff!important;padding:24px!important;overflow:visible!important;box-shadow:none!important;margin:0!important;left:0!important;display:block!important;position:relative!important;";
            }
          }
        },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' },
        pagebreak:    { mode: ['css', 'legacy'] }
      };
    }

    async function downloadPDF(event) {
      const btn = event ? event.currentTarget : null;
      if (btn) {
        btn.disabled = true;
        btn.innerText = "⏳ Generando...";
      }
      
      const element = document.getElementById("pdf-report-content");
      if (!element) return;
      
      const originalTitle = document.title;
      const filename = `Resumen_Comercial_${new Date().toISOString().split('T')[0]}.pdf`;
      document.title = filename;
      
      const appContainer = document.querySelector(".app-container");
      const parent = element.parentElement;
      const nextSibling = element.nextSibling;
      
      const originalScrollX = window.scrollX;
      const originalScrollY = window.scrollY;
      window.scrollTo(0, 0);
      
      if (appContainer) appContainer.style.display = "none";
      
      document.body.appendChild(element);
      document.body.classList.add("pdf-generating");
      element.classList.add("pdf-mode");
      element.style.cssText = 'width:960px!important;max-width:960px!important;min-width:960px!important;background:#ffffff!important;padding:24px!important;overflow:visible!important;box-shadow:none!important;margin:0!important;left:0!important;display:block!important;position:relative!important;';
      
      await new Promise(resolve => setTimeout(resolve, 150));
      
      try {
        const opt = getPDFOptions(filename);
        const pdfBlob = await html2pdf().set(opt).from(element).output('blob');
        
        // 1. Guardar el PDF en el servidor
        const urlParams = new URLSearchParams(window.location.search);
        const reportDate = urlParams.get('date') || new Date().toISOString().split('T')[0];
        const formData = new FormData();
        formData.append("pdf", pdfBlob, filename);
        
        const response = await fetch(`/save-pdf?date=${reportDate}`, {
          method: "POST",
          body: formData
        });
        const result = await response.json();
        
        if (result.success) {
          // 2. Iniciar descarga 100% segura solicitando el archivo al servidor.
          // El servidor devuelve 'Content-Disposition: attachment', por lo que no se cambia de página
          // y el navegador respeta el nombre del archivo de forma nativa.
          window.location.href = `/download-pdf?date=${reportDate}`;
          
          setTimeout(() => {
            alert(`✅ PDF generado y descargado con éxito.\nCopia de seguridad guardada en el servidor.`);
          }, 500);
        } else {
          alert(`⚠️ Error al guardar PDF en el servidor local: ${result.error}`);
        }
      } catch (err) {
        console.error(err);
        alert("Error al generar PDF");
      } finally {
        element.classList.remove("pdf-mode");
        element.style.cssText = '';
        document.body.classList.remove("pdf-generating");
        
        if (nextSibling) {
          parent.insertBefore(element, nextSibling);
        } else {
          parent.appendChild(element);
        }
        
        if (appContainer) appContainer.style.display = "";
        window.scrollTo(originalScrollX, originalScrollY);
        document.title = originalTitle;
        if (btn) {
          btn.disabled = false;
          btn.innerText = "📄 Descargar PDF";
        }
      }
    }

    function openEmailModal() {
      document.getElementById("email-modal").style.display = "flex";
      document.getElementById("dest-email").focus();
    }
    
    function closeEmailModal() {
      document.getElementById("email-modal").style.display = "none";
    }

    async function sendEmail(event) {
      event.preventDefault();
      const emailInput = document.getElementById("dest-email");
      const toEmail = emailInput.value.trim();
      const sendBtn = document.getElementById("send-email-btn");
      
      if (!toEmail) return;
      
      const element = document.getElementById("pdf-report-content");
      if (!element) return;
      
      sendBtn.disabled = true;
      sendBtn.innerText = "⏳ Generando PDF...";
      
      const parent = element.parentElement;
      const nextSibling = element.nextSibling;
      const appContainer = document.querySelector(".app-container");
      
      // Guardar posición de scroll actual y desplazar al origen
      const originalScrollX = window.scrollX;
      const originalScrollY = window.scrollY;
      window.scrollTo(0, 0);
      
      // Ocultar temporalmente el contenedor principal de la aplicación
      if (appContainer) appContainer.style.display = "none";
      
      // Mover el reporte directamente al body
      document.body.appendChild(element);
      
      // Aplicar clases y estilos estáticos limpios
      document.body.classList.add("pdf-generating");
      element.classList.add("pdf-mode");
      element.style.cssText = 'width:960px!important;max-width:960px!important;min-width:960px!important;background:#ffffff!important;padding:24px!important;overflow:visible!important;box-shadow:none!important;margin:0!important;left:0!important;display:block!important;position:relative!important;';
      
      // Esperar a que el navegador complete el reflow
      await new Promise(resolve => setTimeout(resolve, 150));
      
      try {
        const filename = `Resumen_Comercial_${new Date().toISOString().split('T')[0]}.pdf`;
        const opt = getPDFOptions(filename);
        
        // Generar PDF como blob
        const pdfBlob = await html2pdf().set(opt).from(element).output('blob');
        
        // Desactivar clases y limpiar estilos inline
        element.classList.remove("pdf-mode");
        element.style.cssText = '';
        document.body.classList.remove("pdf-generating");
        
        // Devolver a su posición original en el DOM
        if (nextSibling) {
          parent.insertBefore(element, nextSibling);
        } else {
          parent.appendChild(element);
        }
        
        // Mostrar de nuevo el contenedor principal
        if (appContainer) appContainer.style.display = "";
        
        // Restaurar posición de scroll original
        window.scrollTo(originalScrollX, originalScrollY);
        
        sendBtn.innerText = "📧 Enviando Correo...";
        
        // Preparar FormData
        const formData = new FormData();
        formData.append("to_email", toEmail);
        formData.append("pdf", pdfBlob, filename);
        
        // Enviar al servidor
        const response = await fetch("/send-email", {
          method: "POST",
          body: formData
        });
        
        const result = await response.json();
        if (result.success) {
          alert("✅ ¡Resumen enviado con éxito por correo electrónico!");
          closeEmailModal();
        } else {
          alert("❌ Error: " + result.error);
        }
      } catch (err) {
        // Desactivar clases y limpiar estilos inline en caso de error
        element.classList.remove("pdf-mode");
        element.style.cssText = '';
        document.body.classList.remove("pdf-generating");
        
        // Devolver a su posición original en el DOM
        if (nextSibling) {
          parent.insertBefore(element, nextSibling);
        } else {
          parent.appendChild(element);
        }
        
        // Mostrar de nuevo el contenedor principal
        if (appContainer) appContainer.style.display = "";
        
        // Restaurar posición de scroll original
        window.scrollTo(originalScrollX, originalScrollY);
        
        console.error("Error al enviar email:", err);
        alert("❌ Ocurrió un error al enviar el correo.");
      } finally {
        sendBtn.disabled = false;
        sendBtn.innerText = "📤 Enviar Correo";
      }
    }
    function switchTab(tabId) {
      // Ocultar todos los contenidos de pestaña
      document.querySelectorAll(".tab-content").forEach(el => {
        el.classList.remove("active");
      });
      // Mostrar el contenido de la pestaña seleccionada
      const targetContent = document.getElementById(tabId);
      if (targetContent) {
        targetContent.classList.add("active");
      }
      
      // Quitar clase activa de todos los botones
      document.querySelectorAll(".nav-item").forEach(el => {
        el.classList.remove("active");
      });
      // Añadir clase activa al botón correspondiente
      const targetNav = document.getElementById("nav-" + tabId);
      if (targetNav) {
        targetNav.classList.add("active");
      }
      
      // Guardar en localStorage
      localStorage.setItem("selectedTab", tabId);
    }
    
    document.addEventListener("DOMContentLoaded", () => {
      // Cargar pestaña persistida o por defecto
      const savedTab = localStorage.getItem("selectedTab");
      const validTabs = ["resumen-ejecutivo", "resumen-diario", "previsiones-cierre", "calendario-entregas", "alertas-auditoria", "cartera-comparativas", "importar-datos"];
      if (savedTab && validTabs.includes(savedTab) && document.getElementById(savedTab)) {
        switchTab(savedTab);
      } else {
        switchTab("resumen-ejecutivo");
      }

      // Smooth progress bar animations
      setTimeout(() => {
        document.querySelectorAll(".bar-fill").forEach(fill => {
          const styleWidth = fill.style.width;
          fill.style.width = '0%';
          setTimeout(() => { fill.style.width = styleWidth; }, 50);
        });
        document.querySelectorAll(".budget-fill").forEach(fill => {
          const styleWidth = fill.style.width;
          fill.style.width = '0%';
          setTimeout(() => { fill.style.width = styleWidth; }, 50);
        });
      }, 100);

      // Habilitar ordenación en todas las tablas
      function makeTablesSortable() {
        const tables = document.querySelectorAll("table");
        tables.forEach(table => {
          // Normalizar estructura de la tabla si no tiene thead/tbody explícitos
          let thead = table.querySelector("thead");
          let tbody = table.querySelector("tbody");
          
          if (!thead) {
            thead = document.createElement("thead");
            const firstRow = table.querySelector("tr");
            if (firstRow) {
              const cells = firstRow.querySelectorAll("td, th");
              const newHeaderRow = document.createElement("tr");
              cells.forEach(cell => {
                const th = document.createElement("th");
                th.innerHTML = cell.innerHTML;
                for (let attr of cell.attributes) {
                  th.setAttribute(attr.name, attr.value);
                }
                newHeaderRow.appendChild(th);
              });
              thead.appendChild(newHeaderRow);
              firstRow.remove();
              table.insertBefore(thead, table.firstChild);
            }
          }
          
          if (!tbody) {
            tbody = document.createElement("tbody");
            const remainingRows = Array.from(table.querySelectorAll("tr")).filter(row => !thead.contains(row));
            remainingRows.forEach(row => tbody.appendChild(row));
            table.appendChild(tbody);
          }
          
          const headers = thead.querySelectorAll("th");
          if (!headers.length) return;
          
          table.setAttribute("data-sort-col", "-1");
          table.setAttribute("data-sort-dir", "asc");
          
          headers.forEach((header, index) => {
            header.style.cursor = "pointer";
            header.style.userSelect = "none";
            header.classList.add("sortable-header");
            
            // Eliminar indicador previo si existiera
            const prevInd = header.querySelector(".sort-indicator");
            if (prevInd) prevInd.remove();
            
            const indicator = document.createElement("span");
            indicator.className = "sort-indicator";
            indicator.innerHTML = "↕";
            header.appendChild(indicator);
            
            header.addEventListener("click", () => {
              const currentSortCol = parseInt(table.getAttribute("data-sort-col") || "-1");
              let dir = table.getAttribute("data-sort-dir") || "asc";
              
              if (currentSortCol === index) {
                dir = dir === "asc" ? "desc" : "asc";
              } else {
                dir = "asc";
              }
              
              table.setAttribute("data-sort-col", index.toString());
              table.setAttribute("data-sort-dir", dir);
              
              headers.forEach((h, idx) => {
                const ind = h.querySelector(".sort-indicator");
                if (ind) {
                  if (idx === index) {
                    ind.innerHTML = dir === "asc" ? "▲" : "▼";
                    ind.classList.add("active");
                  } else {
                    ind.innerHTML = "↕";
                    ind.classList.remove("active");
                  }
                }
              });
              
              const rows = Array.from(tbody.querySelectorAll("tr"));
              if (rows.length === 0) return;
              
              // Determinar tipo de columna por sus valores no vacíos
              const values = rows.map(row => {
                const cell = row.cells[index];
                return cell ? cell.textContent.trim() : "";
              });
              
              const parseDate = (str) => {
                const match = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
                if (match) {
                  return new Date(match[3], match[2] - 1, match[1]).getTime();
                }
                return null;
              };
              
              const parseNum = (str) => {
                let cleanStr = str.replace(/[€$£%\sA-Za-z]/g, "");
                if (cleanStr.includes(",") && cleanStr.includes(".")) {
                  cleanStr = cleanStr.replace(/\./g, "").replace(/,/g, ".");
                } else if (cleanStr.includes(",")) {
                  cleanStr = cleanStr.replace(/,/g, ".");
                } else if (cleanStr.includes(".")) {
                  const parts = cleanStr.split(".");
                  if (parts.length > 2 || (parts.length === 2 && parts[1].length === 3)) {
                    cleanStr = cleanStr.replace(/\./g, "");
                  }
                }
                const val = parseFloat(cleanStr);
                return isNaN(val) ? -Infinity : val;
              };
              
              let isDateCol = true;
              let isNumCol = true;
              let checkedCount = 0;
              
              for (let val of values) {
                if (!val || val === "-" || val === "Sin fecha" || val.toLowerCase() === "n/d") continue;
                checkedCount++;
                
                // Fecha
                const firstWord = val.split(" ")[0];
                if (!/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(firstWord)) {
                  isDateCol = false;
                }
                
                // Número / importe / porcentaje
                let cleanStr = val.replace(/[€$£%\sA-Za-z]/g, "");
                if (cleanStr.includes(",") && cleanStr.includes(".")) {
                  cleanStr = cleanStr.replace(/\./g, "").replace(/,/g, ".");
                } else if (cleanStr.includes(",")) {
                  cleanStr = cleanStr.replace(/,/g, ".");
                }
                if (!/\d/.test(val) || isNaN(parseFloat(cleanStr))) {
                  isNumCol = false;
                }
              }
              
              if (checkedCount === 0) {
                isDateCol = false;
                isNumCol = false;
              }
              
              rows.sort((a, b) => {
                let cellA = a.cells[index];
                let cellB = b.cells[index];
                let valA = cellA ? cellA.textContent.trim() : "";
                let valB = cellB ? cellB.textContent.trim() : "";
                
                const isValAEmpty = valA === "" || valA === "-" || valA.toLowerCase() === "n/d";
                const isValBEmpty = valB === "" || valB === "-" || valB.toLowerCase() === "n/d";
                if (isValAEmpty && isValBEmpty) return 0;
                if (isValAEmpty) return 1;
                if (isValBEmpty) return -1;
                
                if (isDateCol) {
                  const dateA = parseDate(valA.split(" ")[0]);
                  const dateB = parseDate(valB.split(" ")[0]);
                  if (dateA !== null && dateB !== null) {
                    return dir === "asc" ? dateA - dateB : dateB - dateA;
                  }
                }
                
                if (isNumCol) {
                  const numA = parseNum(valA);
                  const numB = parseNum(valB);
                  return dir === "asc" ? numA - numB : numB - numA;
                }
                
                return dir === "asc" 
                  ? valA.localeCompare(valB, 'es', { numeric: true, sensitivity: 'base' })
                  : valB.localeCompare(valA, 'es', { numeric: true, sensitivity: 'base' });
              });
              
              rows.forEach(row => tbody.appendChild(row));
            });
          });
        });
      }
      
      makeTablesSortable();
    });
  