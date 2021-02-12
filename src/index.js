var XLSX = require("xlsx");
var schema = {};
var modelo = {};

const ExcelAJSON = () => {
  const excel = XLSX.readFile(
    "C:\\Users\\sanqu\\Desktop\\ExcelToJSON\\Inventario-cursos-revisados.xlsx"
  );
  var nombreHoja = excel.SheetNames;
  let datos = XLSX.utils.sheet_to_json(excel.Sheets[nombreHoja[0]]);
  const jDatos = [];
  for (let i = 0; i < datos.length; i++) {
    var moment = require("moment"); // require
    const dato = datos[i];
    schema[i] = {
      groups: [
        {
          legend: "Rúbrica de valoración - Estado de cursos",
          styleClasses: "html2pdf__page-break",
          fields: [
            {
              type: "input",
              inputType: "text",
              label: "Nombre del curso",
              placeholder: "Nombre del curso",
              required: true,
              model: "course_name",
              inputName: "course_name",
            },
            {
              type: "input",
              inputType: "text",
              label: "Unidad Académica a la que pertenece",
              placeholder: "Unidad Académica a la que pertenece",
              required: true,
              model: "academic_unit",
              inputName: "academic_unit",
            },
            {
              type: "input",
              inputType: "text",
              label: "Código (En MARES)",
              placeholder: "Código (En MARES)",
              required: true,
              model: "mares_code",
              inputName: "mares_code",
            },
            {
              type: "input",
              inputType: "text",
              label: "URL",
              placeholder: "URL",
              required: true,
              model: "url",
              inputName: "url",
            },
            {
              type: "input",
              inputType: "text",
              label: "Docente a cargo",
              placeholder: "Docente a cargo",
              required: true,
              model: "teacher",
              inputName: "teacher",
            },
            {
              type: "input",
              inputType: "text",
              label: "Id del curso",
              placeholder: "Id del curso",
              model: "course_id",
              inputName: "course_id",
            },
            {
              type: "select",
              label: "Plataforma",
              model: "platform",
              values: [
                "Udearroba Internos",
                "Udearroba Internos",
                "Aprendeenlinea principal",
                "Aprendeenlinea investigación",
              ],
            },
          ],
        },
        {
          legend: "Generalidades",
          styleClasses: "html2pdf__page-break",
          fields: [
            {
              type: "textArea",
              label:
                "Disposición de elementos de la plataforma: A continuación, haga una breve descripción general de cómo están distribuidos los recursos y herramientas en el aula virtual. Se trata de dar una percepción respecto a qué tan intuitivo resulta para los estudiantes su navegación, y cómo incide ello en su proceso de formación académica.",
              model: "disposition_elements",
              placeholder: "Disposicion de Elementos",
            },
            {
              type: "matrix",
              label: "Criterios para generalidades",
              model: "general_criteria",
              required: true,
              questions: [
                {
                  name: "Cuenta con el mismo título en MARES y la plataforma",
                  id: "has_same_amres_title",
                  required: true,
                },
                {
                  name:
                    "Posee un Texto o video de bienvenida que ubique al estudiante, explicando de forma concisa en qué consistirá el curso.",
                  id: "has_welcome_text",
                  required: true,
                },
                {
                  name:
                    "Presenta una Ficha del Tutor que es la Hoja de vida resumida del profesor.",
                  id: "report_file",
                  required: true,
                },
                {
                  name:
                    "Cuenta con una metodología donde haga una descripción y explicación del qué, cómo y para qué del curso, enunciando el objetivo, temáticas, unidades o módulos, medios de comunicación, estrategias didácticas y evaluativas de todo el curso.",
                  id: "has_methodology",
                  required: true,
                },
                {
                  name:
                    "Establece un cronograma que dé claridad en las semanas de estudio, las temáticas, las actividades a desarrollar, encuentros sincrónicos y la evaluación.",
                  id: "set_schedule",
                  required: true,
                },
                {
                  name:
                    "Presenta la evaluación dónde se muestran los porcentajes evaluativos, además, de incluir rúbricas para la evaluación de las actividades.",
                  id: "report_test",
                  required: true,
                },
                {
                  name:
                    "Contiene un Mapa conceptual u organigrama del curso dónde de relación lógica entre los diferentes conceptos del curso.",
                  id: "has_conceptual_map",
                  required: true,
                },
                {
                  name: "Dispone el Programa del curso de forma descargable",
                  id: "downloadable_course",
                  required: true,
                },
                {
                  name:
                    "Cuenta con foros de Novedades, Dudas e inquietudes y presentación.",
                  id: "has_news_forum",
                  required: true,
                },
              ],
              values: [
                {
                  name: "Cumple",
                  value: "1",
                },
                {
                  name: "Cumple parcialmente",
                  value: "2",
                },
                {
                  name: "No cumple",
                  value: "3",
                },
                {
                  name: "No Aplica",
                  value: "4",
                },
              ],
            },
            {
              type: "textArea",
              label: "Observación para generalidades",
              model: "general_observation",
              help:
                "Solo en caso de seleccionar Cumple parcialmente o No cumple",
              placeholder: "Observación para generalidades",
            },
          ],
        },
        {
          legend: "Sección 1: Unidades o Módulos",
          styleClasses: "html2pdf__page-break",
          id: "0",
          fields: [
            {
              type: "input",
              inputType: "text",
              label: "Nombre del módulo / unidad / sección",
              placeholder: "Nombre del módulo / unidad / sección",
              model: "unity_name",
            },
            {
              type: "textArea",
              label:
                "Disposición de elementos en la plataforma: A continuación, haga una breve descripción general de cómo están distribuidos los recursos y herramientas en el aula virtual. Se trata de dar una percepción respecto a qué tan intuitivo resulta para los estudiantes su navegación, y cómo incide ello en su proceso de formación académica.",
              model: "disposition_elements_unity",
              placeholder: "Disposicion de Elementos",
            },
            {
              type: "matrix",
              label: "Criterios para la Sección",
              model: "section_criteria",
              required: true,
              questions: [
                {
                  name:
                    "Cada Unidad o Módulo cuenta con una introducción de máximo 2 párrafos donde se expliquen los ejes temáticos que se abordarán.",
                  id: "has_a_instroduction",
                  required: true,
                },
                {
                  name:
                    "Dispone en cada módulo o unidad los objetivos que se pretenden cumplir.",
                  id: "has_a_objectives",
                  required: true,
                },
                {
                  name:
                    "Cada unidad módulo cuenta con Material Fundamental (material realizado por el docente)",
                  id: "has_a_basic_material",
                  required: true,
                },
                {
                  name:
                    "Se dispone el material de apoyo con correcta citación (normas APA) y reconocimiento de créditos.",
                  id: "has_a_support_material",
                  required: true,
                },
                {
                  name:
                    "Presenta una Guía de estudio donde se explica en cada unidad o módulo lo que el estudiante debe realizar, se aconseja que tenga un paso a paso.",
                  id: "has_a_study_guide",
                  required: true,
                },
                {
                  name:
                    "Dispone un cuestionario de Autoevaluación ya sea en cada unidad o módulo o al inicio y al finalizar el curso.",
                  id: "has_a_module_test",
                  required: true,
                },
                {
                  name:
                    "La autoevaluación cuenta con Una introducción donde se le cuente al estudiante cuál es el fin de este ejercicio y una serie de preguntas que correspondan a lo trabajado en la unidad o módulo y den cuenta de lo aprendido.",
                  id: "has_a_instroduction_test",
                  required: true,
                },
                {
                  name:
                    "Las Actividades cuentan con: título, modalidad (individual grupal), producto esperado, recursos y materiales que el estudiante debe abordar para el desarrollo de la actividad.",
                  id: "has_a_activity_title",
                  required: true,
                },
                {
                  name:
                    "Las indicaciones de las actividades son claras y concisas, con criterios establecidos para la evaluación de los productos.",
                  id: "has_a_activity_indications",
                  required: true,
                },
                {
                  name:
                    "Son claras la fechas de entrega de las actividades y el espacio por el cual se recibirán.",
                  id: "has_a_activity_date",
                  required: true,
                },
              ],
              values: [
                {
                  name: "Cumple",
                  value: "1",
                },
                {
                  name: "Cumple parcialmente",
                  value: "2",
                },
                {
                  name: "No cumple",
                  value: "3",
                },
                {
                  name: "No Aplica",
                  value: "4",
                },
              ],
            },
            {
              type: "textArea",
              label: "Observación para seccion",
              help:
                "Solo en caso de seleccionar Cumple parcialmente o No cumple",
              model: "section_observation",
              placeholder: "Disposicionn de Elementos",
            },
          ],
        },
      ],
    };
    modelo[i] = {
      course_name: dato["Nombre del curso"],
      academic_unit: dato["Unidad Académica a la que pertenece"],
      mares_code: dato["Código (En MARES)"],
      url: dato["URL"],
      teacher: dato["Docente a cargo"],
      course_id: "1",
      // course_id: dato["Id del curso"],
      platform: dato["Plataforma"],
      disposition_elements:
        dato[
          "Disposición de elementos de la plataforma: A continuación, haga una breve descripción general de cómo están distribuidos los recursos y herramientas en el aula virtual. Se trata de dar una percepción respecto a qué tan intuitivo resulta para los estudiantes su navegación, y cómo incide ello en su proceso de formación académica."
        ],
      general_criteria: {
        has_same_amres_title: matrixValue(
          dato[
            "Criterios para generalidades [Cuenta con el mismo título en MARES y la plataforma]"
          ]
        ),
        has_welcome_text: matrixValue(
          dato[
            "Criterios para generalidades [Posee un Texto o video de bienvenida que  ubique al estudiante, explicando de forma concisa en qué consistirá el curso.]"
          ]
        ),
        report_file: matrixValue(
          dato[
            "Criterios para generalidades [Presenta una Ficha del Tutor que es la Hoja de vida resumida del profesor.]"
          ]
        ),
        has_methodology: matrixValue(
          dato[
            "Criterios para generalidades [Cuenta con una metodología donde haga una descripción y explicación del qué, cómo y para qué del curso, enunciando el objetivo, temáticas, unidades o módulos, medios de comunicación, estrategias didácticas y evaluativas de todo el curso.]"
          ]
        ),
        set_schedule: matrixValue(
          dato[
            "Criterios para generalidades [Establece un cronograma que dé claridad en las semanas de estudio, las temáticas, las actividades a desarrollar, encuentros sincrónicos y la evaluación.]"
          ]
        ),
        report_test: matrixValue(
          dato[
            "Criterios para generalidades [Presenta la evaluación dónde se muestran los porcentajes evaluativos, además, de incluir rúbricas para la evaluación de las actividades.]"
          ]
        ),
        has_conceptual_map: matrixValue(
          dato[
            "Criterios para generalidades [Contiene un Mapa conceptual u organigrama del curso dónde de relación lógica entre los diferentes conceptos del curso.]"
          ]
        ),
        downloadable_course: matrixValue(
          dato[
            "Criterios para generalidades [Dispone el Programa del curso de forma descargable]"
          ]
        ),
        has_news_forum: matrixValue(
          dato[
            "Criterios para generalidades [Cuenta con foros de Novedades, Dudas e inquietudes y presentación.]"
          ]
        ),
      },
      general_observation: dato["Observación para generalidades"],
      unity_name: dato["Nombre del módulo / unidad / sección 1"],
      disposition_elements_unity:
        dato[
          "Disposición de elementos en la plataforma: A continuación, haga una breve descripción general de cómo están distribuidos los recursos y herramientas en el aula virtual. Se trata de dar una percepción respecto a qué tan intuitivo resulta para los estudiantes su navegación, y cómo incide ello en su proceso de formación académica."
        ],
      section_criteria: {
        has_a_instroduction: matrixValue(
          dato[
            "Criterios para sección 1 [Cada Unidad o Módulo cuenta con una  introducción de máximo 2 párrafos  donde se expliquen los ejes temáticos que se abordarán.]"
          ]
        ),
        has_a_objectives: matrixValue(
          dato[
            "Criterios para sección 1 [Dispone en cada módulo o unidad los objetivos que se pretenden cumplir.]"
          ]
        ),
        has_a_basic_material: matrixValue(
          dato[
            "Criterios para sección 1 [Cada unidad módulo cuenta con Material Fundamental (material realizado  por el docente)]"
          ]
        ),
        has_a_support_material: matrixValue(
          dato[
            "Criterios para sección 1 [Se dispone el material de apoyo con correcta citación (normas APA) y reconocimiento de créditos.]"
          ]
        ),
        has_a_study_guide: matrixValue(
          dato[
            "Criterios para sección 1 [Presenta una Guía de estudio donde se explica en cada unidad o módulo lo que el estudiante debe realizar, se aconseja que tenga un paso  a paso.]"
          ]
        ),
        has_a_module_test: matrixValue(
          dato[
            "Criterios para sección 1 [Dispone un cuestionario de Autoevaluación ya sea  en cada unidad o módulo o al inicio y al finalizar el curso.]"
          ]
        ),
        has_a_instroduction_test: matrixValue(
          dato[
            "Criterios para sección 1 [La autoevaluación cuenta con Una introducción donde se le cuente al estudiante cuál es el fin de este ejercicio y una serie de preguntas que  correspondan a lo trabajado en la unidad o módulo y den cuenta de lo aprendido.]"
          ]
        ),
        has_a_activity_title: matrixValue(
          dato[
            "Criterios para sección 1 [Las Actividades cuentan con: título, modalidad (individual grupal), producto esperado, recursos y materiales que el estudiante debe abordar para el desarrollo de la actividad.]"
          ]
        ),
        has_a_activity_indications: matrixValue(
          dato[
            "Criterios para sección 1 [Las indicaciones de las actividades son claras y concisas, con criterios establecidos para la evaluación de los productos.]"
          ]
        ),
        has_a_activity_date: matrixValue(
          dato[
            "Criterios para sección 1 [Son claras la fechas de entrega de las actividades y el espacio por el cual se recibirán.]"
          ]
        ),
      },
      section_observation: dato["Observación para seccion 1"],
    };

    var date = new Date((dato["Marca temporal"] - (25567 + 2)) * 86400 * 1000);
    var localTime = new Date(
      date.getTime() + new Date().getTimezoneOffset() * 60000
    );

    jDatos.push({
      ...dato,
      "Marca temporal": moment(localTime).format("YYYY-MM-DD hh:mm:ss"),
    });

    const number = Object.keys(jDatos[i])[
      Object.keys(jDatos[i]).length - 1
    ].replace("Nombre del módulo / unidad / sección ", "");

    if (!modelo[i].unity_name) {
      modelo[i].unity_name = "Sección 1";
    }

    if (!isNaN(number)) {
      var idSeccion = 1;
      while (idSeccion < number) {
        var group = {
          legend: "Sección 1: Unidades o Módulos",
          styleClasses: "html2pdf__page-break",
          id: "0",
          fields: [
            {
              type: "input",
              inputType: "text",
              label: "Nombre del módulo / unidad / sección",
              placeholder: "Nombre del módulo / unidad / sección",
              model: "unity_name",
            },
            {
              type: "textArea",
              label:
                "Disposición de elementos en la plataforma: A continuación, haga una breve descripción general de cómo están distribuidos los recursos y herramientas en el aula virtual. Se trata de dar una percepción respecto a qué tan intuitivo resulta para los estudiantes su navegación, y cómo incide ello en su proceso de formación académica.",
              model: "disposition_elements_unity",
              placeholder: "Disposicion de Elementos",
            },
            {
              type: "matrix",
              label: "Criterios para la Sección",
              model: "section_criteria",
              required: true,
              questions: [
                {
                  name:
                    "Cada Unidad o Módulo cuenta con una introducción de máximo 2 párrafos donde se expliquen los ejes temáticos que se abordarán.",
                  id: "has_a_instroduction",
                  required: true,
                },
                {
                  name:
                    "Dispone en cada módulo o unidad los objetivos que se pretenden cumplir.",
                  id: "has_a_objectives",
                  required: true,
                },
                {
                  name:
                    "Cada unidad módulo cuenta con Material Fundamental (material realizado por el docente)",
                  id: "has_a_basic_material",
                  required: true,
                },
                {
                  name:
                    "Se dispone el material de apoyo con correcta citación (normas APA) y reconocimiento de créditos.",
                  id: "has_a_support_material",
                  required: true,
                },
                {
                  name:
                    "Presenta una Guía de estudio donde se explica en cada unidad o módulo lo que el estudiante debe realizar, se aconseja que tenga un paso a paso.",
                  id: "has_a_study_guide",
                  required: true,
                },
                {
                  name:
                    "Dispone un cuestionario de Autoevaluación ya sea en cada unidad o módulo o al inicio y al finalizar el curso.",
                  id: "has_a_module_test",
                  required: true,
                },
                {
                  name:
                    "La autoevaluación cuenta con Una introducción donde se le cuente al estudiante cuál es el fin de este ejercicio y una serie de preguntas que correspondan a lo trabajado en la unidad o módulo y den cuenta de lo aprendido.",
                  id: "has_a_instroduction_test",
                  required: true,
                },
                {
                  name:
                    "Las Actividades cuentan con: título, modalidad (individual grupal), producto esperado, recursos y materiales que el estudiante debe abordar para el desarrollo de la actividad.",
                  id: "has_a_activity_title",
                  required: true,
                },
                {
                  name:
                    "Las indicaciones de las actividades son claras y concisas, con criterios establecidos para la evaluación de los productos.",
                  id: "has_a_activity_indications",
                  required: true,
                },
                {
                  name:
                    "Son claras la fechas de entrega de las actividades y el espacio por el cual se recibirán.",
                  id: "has_a_activity_date",
                  required: true,
                },
              ],
              values: [
                {
                  name: "Cumple",
                  value: "1",
                },
                {
                  name: "Cumple parcialmente",
                  value: "2",
                },
                {
                  name: "No cumple",
                  value: "3",
                },
                {
                  name: "No Aplica",
                  value: "4",
                },
              ],
            },
            {
              type: "textArea",
              label: "Observación para seccion",
              help:
                "Solo en caso de seleccionar Cumple parcialmente o No cumple",
              model: "section_observation",
              placeholder: "Disposicionn de Elementos",
            },
          ],
        };
        group.id = idSeccion;
        group.legend = "Sección " + (idSeccion + 1) + ": Unidades o Módulos";
        sectionModel = {};
        sectionModel["unity_name" + idSeccion] =
          dato["Nombre del módulo / unidad / sección " + (idSeccion + 1)];
        sectionModel["disposition_elements_unity" + idSeccion] =
          dato[
            "Disposición de elementos en la plataforma: A continuación, haga una breve descripción general de cómo están distribuidos los recursos y herramientas en el aula virtual. Se trata de dar una percepción respecto a qué tan intuitivo resulta para los estudiantes su navegación, y cómo incide ello en su proceso de formación académica."
          ];
        sectionModel["section_observation" + idSeccion] =
          dato["Observación para seccion " + (idSeccion + 1)];
        sectionModel["section_criteria" + idSeccion] = {};
        sectionModel["section_criteria" + idSeccion][
          "has_a_activity_date" + idSeccion
        ] = matrixValue(
          dato[
            "Criterios para sección " +
              (idSeccion + 1) +
              " [Son claras la fechas de entrega de las actividades y el espacio por el cual se recibirán.]"
          ]
        );
        sectionModel["section_criteria" + idSeccion][
          "has_a_activity_indications" + idSeccion
        ] = matrixValue(
          dato[
            "Criterios para sección " +
              (idSeccion + 1) +
              " [Las indicaciones de las actividades son claras y concisas, con criterios establecidos para la evaluación de los productos.]"
          ]
        );
        sectionModel["section_criteria" + idSeccion][
          "has_a_activity_title" + idSeccion
        ] = matrixValue(
          dato[
            "Criterios para sección " +
              (idSeccion + 1) +
              " [Las Actividades cuentan con: título, modalidad (individual grupal), producto esperado, recursos y materiales que el estudiante debe abordar para el desarrollo de la actividad.]"
          ]
        );
        sectionModel["section_criteria" + idSeccion][
          "has_a_module_test" + idSeccion
        ] = matrixValue(
          dato[
            "Criterios para sección " +
              (idSeccion + 1) +
              " [Dispone un cuestionario de Autoevaluación ya sea  en cada unidad o módulo o al inicio y al finalizar el curso.]"
          ]
        );
        sectionModel["section_criteria" + idSeccion][
          "has_a_instroduction_test" + idSeccion
        ] = matrixValue(
          dato[
            "Criterios para sección " +
              (idSeccion + 1) +
              " [La autoevaluación cuenta con Una introducción donde se le cuente al estudiante cuál es el fin de este ejercicio y una serie de preguntas que  correspondan a lo trabajado en la unidad o módulo y den cuenta de lo aprendido.]"
          ]
        );
        sectionModel["section_criteria" + idSeccion][
          "has_a_study_guide" + idSeccion
        ] = matrixValue(
          dato[
            "Criterios para sección " +
              (idSeccion + 1) +
              " [Presenta una Guía de estudio donde se explica en cada unidad o módulo lo que el estudiante debe realizar, se aconseja que tenga un paso  a paso.]"
          ]
        );
        sectionModel["section_criteria" + idSeccion][
          "has_a_support_material" + idSeccion
        ] = matrixValue(
          dato[
            "Criterios para sección " +
              (idSeccion + 1) +
              " [Se dispone el material de apoyo con correcta citación (normas APA) y reconocimiento de créditos.]"
          ]
        );
        sectionModel["section_criteria" + idSeccion][
          "has_a_basic_material" + idSeccion
        ] = matrixValue(
          dato[
            "Criterios para sección " +
              (idSeccion + 1) +
              " [Cada unidad módulo cuenta con Material Fundamental (material realizado  por el docente)]"
          ]
        );
        sectionModel["section_criteria" + idSeccion][
          "has_a_objectives" + idSeccion
        ] = matrixValue(
          dato[
            "Criterios para sección " +
              (idSeccion + 1) +
              " [Dispone en cada módulo o unidad los objetivos que se pretenden cumplir.]"
          ]
        );
        sectionModel["section_criteria" + idSeccion][
          "has_a_instroduction" + idSeccion
        ] = matrixValue(
          dato[
            "Criterios para sección " +
              (idSeccion + 1) +
              " [Cada Unidad o Módulo cuenta con una  introducción de máximo 2 párrafos  donde se expliquen los ejes temáticos que se abordarán.]"
          ]
        );

        for (var p = 0; p < group.fields.length; p++) {
          group.fields[p].model = group.fields[p].model + idSeccion;
          if (group.fields[p].questions) {
            for (var j = 0; j < group.fields[p].questions.length; j++) {
              group.fields[p].questions[j].id =
                group.fields[p].questions[j].id + idSeccion;
            }
          }
        }
        idSeccion++;
        Object.assign(modelo[i], sectionModel);
        schema[i].groups.push(group);
      }
    }
  }

  const alldates = [];
  for (var i = 0; i < jDatos.length; i++) {
    alldates.push({
      formulario_id: "1",
      user_id: "1",
      curso_id: "1",
      respuesta: JSON.stringify(modelo[i]),
      estructura_respuesta: JSON.stringify(schema[i]),
      created_at: jDatos[i]["Marca temporal"],
      updated_at: jDatos[i]["Marca temporal"],
      estado: "Terminado",
      version: "1",
    });
  }
  const fs = require("fs");
  var json2xls = require("json2xls");
  const filename = "Rubricas-diligenciadas-curadurias.xlsx";

  var xls = json2xls(alldates);
  fs.writeFileSync(filename, xls, "binary", (err) => {
    if (err) {
      console.log("writeFileSync :", err);
    }
    console.log(filename + " file is saved!");
  });
};

function matrixValue(value) {
  switch (value) {
    case "Cumple":
      return "1";
    case "Cumple parcialmente":
      return "2";
    case "No cumple":
      return "3";
    case "No Aplica":
      return "4";
  }
}

ExcelAJSON();
