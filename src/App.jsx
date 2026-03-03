import { Fragment, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  DndContext,
  DragOverlay,
  PointerSensor,
  closestCenter,
  useDroppable,
  useSensor,
  useSensors,
} from '@dnd-kit/core';
import {
  SortableContext,
  arrayMove,
  rectSortingStrategy,
  useSortable,
  verticalListSortingStrategy,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { Copy, Download, FolderKanban, Plus, Trash2, Upload } from 'lucide-react';
import PptxGenJS from 'pptxgenjs';
import { v4 as uuidv4 } from 'uuid';

const STORAGE_KEY = 'storylining-board-v1';
const EXPORT_WIDTH = 13.333;
const EXPORT_HEIGHT = 7.5;
const SECTION_COLOR = '#4d217a';
const CARD_BG = '#d7e4ee';
const CARD_BORDER = '#b7c8d4';

function createEmptySlide(overrides = {}) {
  return {
    id: uuidv4(),
    title: 'New Slide',
    bullets: [''],
    ...overrides,
  };
}

function createEmptySection(overrides = {}) {
  return {
    id: uuidv4(),
    title: 'Section Title',
    slides: [createEmptySlide()],
    ...overrides,
  };
}

function createDefaultBoard() {
  return {
    title: 'Storyline Board',
    sections: [
      createEmptySection({
        title: 'Context',
        slides: [
          createEmptySlide({
            title: 'Current Situation',
            bullets: ['What is happening now?', 'Why it matters', ''],
          }),
        ],
      }),
      createEmptySection({
        title: 'Recommendation',
        slides: [
          createEmptySlide({
            title: 'Decision',
            bullets: ['What should be done', 'Expected impact', ''],
          }),
        ],
      }),
    ],
  };
}

function createProject(name = 'Untitled Board') {
  const board = createDefaultBoard();
  const boardTitle = (name || '').trim() || board.title || 'Untitled Board';

  return {
    id: uuidv4(),
    name: boardTitle,
    board: {
      ...board,
      title: boardTitle,
    },
  };
}

function normalizeSlide(slide) {
  const source = slide && typeof slide === 'object' ? slide : {};
  const bullets = Array.isArray(source.bullets)
    ? source.bullets.map((bullet) => (typeof bullet === 'string' ? bullet : String(bullet ?? '')))
    : [];

  return {
    id: typeof source.id === 'string' && source.id.trim() ? source.id : uuidv4(),
    title: typeof source.title === 'string' ? source.title : 'New Slide',
    bullets: bullets.length > 0 ? bullets : [''],
  };
}

function normalizeSection(section) {
  const source = section && typeof section === 'object' ? section : {};
  const slides = Array.isArray(source.slides) ? source.slides.map(normalizeSlide) : [];

  return {
    id: typeof source.id === 'string' && source.id.trim() ? source.id : uuidv4(),
    title: typeof source.title === 'string' ? source.title : 'Section Title',
    slides: slides.length > 0 ? slides : [createEmptySlide()],
  };
}

function normalizeBoard(board, fallbackTitle = 'Untitled Board') {
  const source = board && typeof board === 'object' ? board : {};
  const sections = Array.isArray(source.sections) ? source.sections.map(normalizeSection) : [];
  const title =
    typeof source.title === 'string' && source.title.trim()
      ? source.title.trim()
      : fallbackTitle.trim() || 'Untitled Board';

  return {
    title,
    sections: sections.length > 0 ? sections : [createEmptySection()],
  };
}

function normalizeProjectEntry(project, index) {
  const source = project && typeof project === 'object' ? project : {};
  const fallback =
    typeof source.name === 'string' && source.name.trim() ? source.name.trim() : 'Board ' + (index + 1);
  const board = normalizeBoard(source.board, fallback);

  return {
    id: typeof source.id === 'string' && source.id.trim() ? source.id : uuidv4(),
    name: board.title,
    board,
  };
}

function parseBoardFromImportedJson(parsed, fallbackTitle = 'Imported Board') {
  if (parsed && typeof parsed === 'object') {
    if (parsed.board && typeof parsed.board === 'object' && !Array.isArray(parsed.board)) {
      const sourceName =
        typeof parsed.name === 'string' && parsed.name.trim() ? parsed.name.trim() : fallbackTitle;
      return normalizeBoard(parsed.board, sourceName);
    }

    if (Array.isArray(parsed.projects) && parsed.projects.length > 0) {
      const activeProject =
        parsed.projects.find((project) => project.id === parsed.activeProjectId) ?? parsed.projects[0];
      const sourceName =
        typeof activeProject?.name === 'string' && activeProject.name.trim()
          ? activeProject.name.trim()
          : fallbackTitle;
      return normalizeBoard(activeProject?.board, sourceName);
    }

    if (Array.isArray(parsed.sections)) {
      return normalizeBoard(parsed, fallbackTitle);
    }
  }

  throw new Error('Invalid board JSON format');
}

function loadProjectStore() {
  if (typeof window === 'undefined') {
    const project = createProject('Default Board');
    return { activeProjectId: project.id, projects: [project] };
  }

  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      const project = createProject('Default Board');
      return { activeProjectId: project.id, projects: [project] };
    }

    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.projects) || parsed.projects.length === 0) {
      const project = createProject('Default Board');
      return { activeProjectId: project.id, projects: [project] };
    }

    const projects = parsed.projects.map(normalizeProjectEntry);
    const activeExists = projects.some((project) => project.id === parsed.activeProjectId);

    return {
      activeProjectId: activeExists ? parsed.activeProjectId : projects[0].id,
      projects,
    };
  } catch {
    const project = createProject('Default Board');
    return { activeProjectId: project.id, projects: [project] };
  }
}

function normalizeBulletsDuringEditing(bullets, keepIndex) {
  const emptyIndexes = bullets
    .map((text, index) => ({ text, index }))
    .filter((item) => item.text.trim() === '')
    .map((item) => item.index);

  if (emptyIndexes.length <= 1) {
    return bullets;
  }

  const keepEmpty =
    keepIndex != null && bullets[keepIndex]?.trim() === '' ? keepIndex : emptyIndexes[0];

  return bullets.filter((text, index) => {
    if (text.trim() !== '') {
      return true;
    }
    return index === keepEmpty;
  });
}

function cleanupBulletsOnCardBlur(bullets) {
  return bullets.filter((text) => text.trim() !== '');
}

function deepCloneSection(section) {
  return {
    ...section,
    id: uuidv4(),
    slides: section.slides.map((slide) => ({
      ...slide,
      id: uuidv4(),
      bullets: [...slide.bullets],
    })),
  };
}

function deepCloneSlide(slide) {
  return {
    ...slide,
    id: uuidv4(),
    bullets: [...slide.bullets],
  };
}

function getSlideTitleKey(slideId) {
  return `slide:${slideId}:title`;
}

function getBulletKey(slideId, bulletIndex) {
  return `slide:${slideId}:bullet:${bulletIndex}`;
}

function getSectionTitleKey(sectionId) {
  return `section:${sectionId}:title`;
}

function safeFileName(input) {
  return input
    .replace(/[<>:"/\\|?*]/g, '')
    .replace(/\s+/g, '_')
    .slice(0, 80);
}
function autoResizeTextarea(node) {
  if (!node) {
    return;
  }

  node.style.height = '0px';
  node.style.height = `${Math.max(node.scrollHeight, 24)}px`;
}


function toSectionLabel(index) {
  let value = index + 1;
  let label = '';

  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
}

function buildSlideNumberMap(sections) {
  const map = {};
  let count = 1;

  sections.forEach((section) => {
    section.slides.forEach((slide) => {
      map[slide.id] = count;
      count += 1;
    });
  });

  return map;
}

function locateSlide(board, slideId) {
  for (let sectionIndex = 0; sectionIndex < board.sections.length; sectionIndex += 1) {
    const section = board.sections[sectionIndex];
    const slideIndex = section.slides.findIndex((slide) => slide.id === slideId);
    if (slideIndex !== -1) {
      return {
        sectionIndex,
        sectionId: section.id,
        slideIndex,
      };
    }
  }

  return null;
}

function resolveOverTarget(over, board) {
  if (!over) {
    return null;
  }

  const data = over.data?.current;
  if (data?.type === 'slide' || data?.type === 'section' || data?.type === 'lane') {
    return data;
  }

  const rawId = String(over.id);

  if (rawId.startsWith('slide:')) {
    const slideId = rawId.slice(6);
    const location = locateSlide(board, slideId);
    if (!location) {
      return null;
    }
    return { type: 'slide', slideId, sectionId: location.sectionId };
  }

  if (rawId.startsWith('section:')) {
    return { type: 'section', sectionId: rawId.slice(8) };
  }

  if (rawId.startsWith('lane:')) {
    return { type: 'lane', sectionId: rawId.slice(5) };
  }

  return null;
}

function SortableSectionRow({ sectionId, children }) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({
    id: `section:${sectionId}`,
    data: { type: 'section', sectionId },
  });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    zIndex: isDragging ? 40 : undefined,
  };

  return children({
    setNodeRef,
    style,
    isDragging,
    dragHandleProps: { ...attributes, ...listeners },
  });
}

function SortableSlideSlot({ slideId, sectionId, children }) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({
    id: `slide:${slideId}`,
    data: { type: 'slide', slideId, sectionId },
  });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    zIndex: isDragging ? 50 : undefined,
    width: 'var(--card-w)',
    flexBasis: 'var(--card-w)',
  };

  return children({
    setNodeRef,
    style,
    isDragging,
    dragHandleProps: { ...attributes, ...listeners },
  });
}

function SlideLaneDroppable({ sectionId, children }) {
  const { setNodeRef, isOver } = useDroppable({
    id: `lane:${sectionId}`,
    data: { type: 'lane', sectionId },
  });

  return children({ setNodeRef, isOver });
}

function App() {
  const [projectStore, setProjectStore] = useState(loadProjectStore);
  const [scaleToFit, setScaleToFit] = useState(true);
  const [pendingFocus, setPendingFocus] = useState(null);
  const [activeDrag, setActiveDrag] = useState(null);
  const inputRefs = useRef(new Map());
  const boardImportInputRef = useRef(null);

  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 5,
      },
    }),
  );

  const activeProject = useMemo(
    () =>
      projectStore.projects.find((project) => project.id === projectStore.activeProjectId) ??
      projectStore.projects[0],
    [projectStore],
  );

  const board = activeProject.board;
  const slideNumberById = useMemo(() => buildSlideNumberMap(board.sections), [board.sections]);

  useEffect(() => {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(projectStore));
  }, [projectStore]);

  useEffect(() => {
    if (!pendingFocus) {
      return;
    }

    const tryFocus = () => {
      const node = inputRefs.current.get(pendingFocus);
      if (!node) {
        return false;
      }
      node.focus();
      if (typeof node.select === 'function') {
        node.select();
      }
      setPendingFocus(null);
      return true;
    };

    if (tryFocus()) {
      return;
    }

    const id = window.requestAnimationFrame(() => {
      tryFocus();
    });

    return () => window.cancelAnimationFrame(id);
  }, [pendingFocus, projectStore]);

  const registerInputRef = useCallback(
    (key) => (node) => {
      if (node) {
        inputRefs.current.set(key, node);
        if (node.tagName === 'TEXTAREA') {
          autoResizeTextarea(node);
        }
      } else {
        inputRefs.current.delete(key);
      }
    },
    [],
  );

  const handleTextEditorInput = useCallback((event) => {
    autoResizeTextarea(event.currentTarget);
  }, []);

  useEffect(() => {
    const id = window.requestAnimationFrame(() => {
      inputRefs.current.forEach((node) => {
        if (node?.tagName === 'TEXTAREA') {
          autoResizeTextarea(node);
        }
      });
    });

    return () => window.cancelAnimationFrame(id);
  }, [board, activeProject.id]);

  const updateActiveBoard = useCallback((updater) => {
    setProjectStore((prev) => {
      const projectIndex = prev.projects.findIndex((project) => project.id === prev.activeProjectId);

      if (projectIndex === -1) {
        return prev;
      }

      const projects = [...prev.projects];
      const current = projects[projectIndex];
      const nextBoard = updater(current.board);
      projects[projectIndex] = { ...current, board: nextBoard };

      return { ...prev, projects };
    });
  }, []);

  const focusSlideTitle = useCallback((slideId) => {
    setPendingFocus(getSlideTitleKey(slideId));
  }, []);

  const focusBullet = useCallback((slideId, bulletIndex) => {
    setPendingFocus(getBulletKey(slideId, bulletIndex));
  }, []);

  const focusSectionTitle = useCallback((sectionId) => {
    setPendingFocus(getSectionTitleKey(sectionId));
  }, []);

  const updateStoryTitle = useCallback((nextTitle) => {
    const normalizedTitle = nextTitle.trim() || 'Untitled Board';

    setProjectStore((prev) => {
      const projectIndex = prev.projects.findIndex((project) => project.id === prev.activeProjectId);
      if (projectIndex === -1) {
        return prev;
      }

      const projects = [...prev.projects];
      const current = projects[projectIndex];
      projects[projectIndex] = {
        ...current,
        name: normalizedTitle,
        board: {
          ...current.board,
          title: normalizedTitle,
        },
      };

      return { ...prev, projects };
    });
  }, []);

  const updateSectionTitle = useCallback(
    (sectionId, nextTitle) => {
      updateActiveBoard((currentBoard) => ({
        ...currentBoard,
        sections: currentBoard.sections.map((section) =>
          section.id === sectionId ? { ...section, title: nextTitle } : section,
        ),
      }));
    },
    [updateActiveBoard],
  );

  const updateSlideTitle = useCallback(
    (sectionId, slideId, nextTitle) => {
      updateActiveBoard((currentBoard) => ({
        ...currentBoard,
        sections: currentBoard.sections.map((section) => {
          if (section.id !== sectionId) {
            return section;
          }

          return {
            ...section,
            slides: section.slides.map((slide) =>
              slide.id === slideId ? { ...slide, title: nextTitle } : slide,
            ),
          };
        }),
      }));
    },
    [updateActiveBoard],
  );

  const updateBulletText = useCallback(
    (sectionId, slideId, bulletIndex, nextText) => {
      updateActiveBoard((currentBoard) => ({
        ...currentBoard,
        sections: currentBoard.sections.map((section) => {
          if (section.id !== sectionId) {
            return section;
          }

          return {
            ...section,
            slides: section.slides.map((slide) => {
              if (slide.id !== slideId) {
                return slide;
              }

              const bullets = slide.bullets.map((text, index) => (index === bulletIndex ? nextText : text));

              return {
                ...slide,
                bullets: normalizeBulletsDuringEditing(bullets, bulletIndex),
              };
            }),
          };
        }),
      }));
    },
    [updateActiveBoard],
  );

  const insertSlideAt = useCallback(
    (sectionId, insertIndex) => {
      const newSlide = createEmptySlide();

      updateActiveBoard((currentBoard) => ({
        ...currentBoard,
        sections: currentBoard.sections.map((section) => {
          if (section.id !== sectionId) {
            return section;
          }

          const slides = [...section.slides];
          const boundedIndex = Math.max(0, Math.min(insertIndex, slides.length));
          slides.splice(boundedIndex, 0, newSlide);
          return { ...section, slides };
        }),
      }));

      focusSlideTitle(newSlide.id);
    },
    [focusSlideTitle, updateActiveBoard],
  );

  const insertSlideAfter = useCallback(
    (sectionId, afterSlideId) => {
      const section = board.sections.find((item) => item.id === sectionId);
      const afterIndex = section ? section.slides.findIndex((slide) => slide.id === afterSlideId) : -1;
      const insertIndex = afterIndex === -1 ? section?.slides.length ?? 0 : afterIndex + 1;
      insertSlideAt(sectionId, insertIndex);
    },
    [board.sections, insertSlideAt],
  );

  const insertSectionAt = useCallback(
    (insertIndex) => {
      const newSection = createEmptySection();

      updateActiveBoard((currentBoard) => {
        const sections = [...currentBoard.sections];
        const boundedIndex = Math.max(0, Math.min(insertIndex, sections.length));
        sections.splice(boundedIndex, 0, newSection);
        return { ...currentBoard, sections };
      });

      focusSectionTitle(newSection.id);
    },
    [focusSectionTitle, updateActiveBoard],
  );

  const duplicateSection = useCallback(
    (sectionId) => {
      let duplicated = null;

      updateActiveBoard((currentBoard) => {
        const sections = [...currentBoard.sections];
        const sectionIndex = sections.findIndex((section) => section.id === sectionId);
        if (sectionIndex === -1) {
          return currentBoard;
        }

        duplicated = deepCloneSection(sections[sectionIndex]);
        sections.splice(sectionIndex + 1, 0, duplicated);
        return { ...currentBoard, sections };
      });

      if (duplicated) {
        focusSectionTitle(duplicated.id);
      }
    },
    [focusSectionTitle, updateActiveBoard],
  );

  const deleteSection = useCallback(
    (sectionId) => {
      let fallbackSectionId = null;

      updateActiveBoard((currentBoard) => {
        const sections = currentBoard.sections.filter((section) => section.id !== sectionId);

        if (sections.length === 0) {
          const replacement = createEmptySection();
          fallbackSectionId = replacement.id;
          return { ...currentBoard, sections: [replacement] };
        }

        fallbackSectionId = sections[0].id;
        return { ...currentBoard, sections };
      });

      if (fallbackSectionId) {
        focusSectionTitle(fallbackSectionId);
      }
    },
    [focusSectionTitle, updateActiveBoard],
  );

  const duplicateSlide = useCallback(
    (sectionId, slideId) => {
      let duplicatedSlide = null;

      updateActiveBoard((currentBoard) => ({
        ...currentBoard,
        sections: currentBoard.sections.map((section) => {
          if (section.id !== sectionId) {
            return section;
          }

          const slides = [...section.slides];
          const slideIndex = slides.findIndex((slide) => slide.id === slideId);
          if (slideIndex === -1) {
            return section;
          }

          duplicatedSlide = deepCloneSlide(slides[slideIndex]);
          slides.splice(slideIndex + 1, 0, duplicatedSlide);
          return { ...section, slides };
        }),
      }));

      if (duplicatedSlide) {
        focusSlideTitle(duplicatedSlide.id);
      }
    },
    [focusSlideTitle, updateActiveBoard],
  );

  const deleteSlide = useCallback(
    (sectionId, slideId) => {
      updateActiveBoard((currentBoard) => ({
        ...currentBoard,
        sections: currentBoard.sections.map((section) => {
          if (section.id !== sectionId) {
            return section;
          }

          const slides = section.slides.filter((slide) => slide.id !== slideId);
          return { ...section, slides };
        }),
      }));

      focusSectionTitle(sectionId);
    },
    [focusSectionTitle, updateActiveBoard],
  );

  const addBulletAfter = useCallback(
    (sectionId, slideId, bulletIndex) => {
      updateActiveBoard((currentBoard) => ({
        ...currentBoard,
        sections: currentBoard.sections.map((section) => {
          if (section.id !== sectionId) {
            return section;
          }

          return {
            ...section,
            slides: section.slides.map((slide) => {
              if (slide.id !== slideId) {
                return slide;
              }

              const bullets = [...slide.bullets];
              bullets.splice(bulletIndex + 1, 0, '');
              return {
                ...slide,
                bullets: normalizeBulletsDuringEditing(bullets, bulletIndex + 1),
              };
            }),
          };
        }),
      }));

      focusBullet(slideId, bulletIndex + 1);
    },
    [focusBullet, updateActiveBoard],
  );

  const deleteBullet = useCallback(
    (sectionId, slideId, bulletIndex) => {
      updateActiveBoard((currentBoard) => ({
        ...currentBoard,
        sections: currentBoard.sections.map((section) => {
          if (section.id !== sectionId) {
            return section;
          }

          return {
            ...section,
            slides: section.slides.map((slide) => {
              if (slide.id !== slideId) {
                return slide;
              }

              const bullets = [...slide.bullets];
              bullets.splice(bulletIndex, 1);
              return { ...slide, bullets };
            }),
          };
        }),
      }));

      if (bulletIndex === 0) {
        focusSlideTitle(slideId);
      } else {
        focusBullet(slideId, bulletIndex - 1);
      }
    },
    [focusBullet, focusSlideTitle, updateActiveBoard],
  );

  const cleanupCardBullets = useCallback(
    (sectionId, slideId) => {
      updateActiveBoard((currentBoard) => ({
        ...currentBoard,
        sections: currentBoard.sections.map((section) => {
          if (section.id !== sectionId) {
            return section;
          }

          return {
            ...section,
            slides: section.slides.map((slide) => {
              if (slide.id !== slideId) {
                return slide;
              }

              return {
                ...slide,
                bullets: cleanupBulletsOnCardBlur(slide.bullets),
              };
            }),
          };
        }),
      }));
    },
    [updateActiveBoard],
  );

  const handleSlideTitleKeyDown = useCallback(
    (event, sectionId, slideId, bulletsCount) => {
      if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
        event.preventDefault();
        insertSlideAfter(sectionId, slideId);
        return;
      }

      if (event.key !== 'Enter') {
        return;
      }

      event.preventDefault();

      if (bulletsCount === 0) {
        updateActiveBoard((currentBoard) => ({
          ...currentBoard,
          sections: currentBoard.sections.map((section) => {
            if (section.id !== sectionId) {
              return section;
            }

            return {
              ...section,
              slides: section.slides.map((slide) =>
                slide.id === slideId ? { ...slide, bullets: [''] } : slide,
              ),
            };
          }),
        }));
      }

      focusBullet(slideId, 0);
    },
    [focusBullet, insertSlideAfter, updateActiveBoard],
  );

  const handleBulletKeyDown = useCallback(
    (event, sectionId, slideId, bulletIndex) => {
      if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
        event.preventDefault();
        insertSlideAfter(sectionId, slideId);
        return;
      }

      if (event.key === 'Enter') {
        event.preventDefault();
        addBulletAfter(sectionId, slideId, bulletIndex);
        return;
      }

      if (event.key === 'Backspace' && event.currentTarget.value.trim() === '') {
        event.preventDefault();
        deleteBullet(sectionId, slideId, bulletIndex);
      }
    },
    [addBulletAfter, deleteBullet, insertSlideAfter],
  );

  const handleBoardSelection = useCallback(
    (selection) => {
      if (selection === '__new_board__') {
        const nextBoardIndex = projectStore.projects.length + 1;
        const defaultName = 'Board ' + nextBoardIndex;
        const name = window.prompt('Board name', defaultName) || defaultName;
        const project = createProject(name.trim() || defaultName);

        setProjectStore((prev) => ({
          activeProjectId: project.id,
          projects: [...prev.projects, project],
        }));
        return;
      }

      setProjectStore((prev) => ({ ...prev, activeProjectId: selection }));
    },
    [projectStore.projects.length],
  );

  const exportBoardJson = useCallback(() => {
    const payload = {
      version: 1,
      exportedAt: new Date().toISOString(),
      board,
    };

    const blob = new Blob([JSON.stringify(payload, null, 2)], {
      type: 'application/json;charset=utf-8',
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = safeFileName(board.title || 'storyline-board') + '.json';
    link.click();
    window.URL.revokeObjectURL(url);
  }, [board]);

  const openBoardImportDialog = useCallback(() => {
    boardImportInputRef.current?.click();
  }, []);

  const importBoardFromJson = useCallback(
    async (event) => {
      const input = event.target;
      const file = input.files?.[0];
      input.value = '';

      if (!file) {
        return;
      }

      try {
        const text = await file.text();
        const parsed = JSON.parse(text);
        const fallbackName =
          file.name.replace(/\.json$/i, '').trim() ||
          'Imported Board ' + (projectStore.projects.length + 1);
        const importedBoard = parseBoardFromImportedJson(parsed, fallbackName);

        const project = {
          id: uuidv4(),
          name: importedBoard.title,
          board: importedBoard,
        };

        setProjectStore((prev) => ({
          activeProjectId: project.id,
          projects: [...prev.projects, project],
        }));
      } catch (error) {
        console.error('Failed to import board JSON:', error);
        window.alert('Unable to import JSON. Please use a valid board export file.');
      }
    },
    [projectStore.projects.length],
  );

  const handleDragStart = useCallback((event) => {
    const activeData = event.active.data.current;
    if (!activeData) {
      return;
    }

    if (activeData.type === 'slide') {
      setActiveDrag({
        type: 'slide',
        slideId: activeData.slideId,
        sectionId: activeData.sectionId,
      });
      return;
    }

    if (activeData.type === 'section') {
      setActiveDrag({
        type: 'section',
        sectionId: activeData.sectionId,
      });
    }
  }, []);

  const handleDragCancel = useCallback(() => {
    setActiveDrag(null);
  }, []);

  const handleDragOver = useCallback(
    ({ active, over }) => {
      const activeData = active.data.current;
      if (!over || activeData?.type !== 'slide') {
        return;
      }

      updateActiveBoard((currentBoard) => {
        const overTarget = resolveOverTarget(over, currentBoard);
        if (!overTarget || (overTarget.type !== 'slide' && overTarget.type !== 'lane')) {
          return currentBoard;
        }

        const source = locateSlide(currentBoard, activeData.slideId);
        if (!source) {
          return currentBoard;
        }

        const destinationSectionId = overTarget.sectionId;
        if (!destinationSectionId || source.sectionId === destinationSectionId) {
          return currentBoard;
        }

        const sections = currentBoard.sections.map((section) => ({
          ...section,
          slides: [...section.slides],
        }));

        const sourceSection = sections[source.sectionIndex];
        const destinationSection = sections.find((section) => section.id === destinationSectionId);

        if (!sourceSection || !destinationSection) {
          return currentBoard;
        }

        const [movedSlide] = sourceSection.slides.splice(source.slideIndex, 1);
        if (!movedSlide) {
          return currentBoard;
        }

        let destinationIndex = destinationSection.slides.length;
        if (overTarget.type === 'slide') {
          const overIndex = destinationSection.slides.findIndex(
            (slide) => slide.id === overTarget.slideId,
          );
          if (overIndex !== -1) {
            destinationIndex = overIndex;
          }
        }

        const boundedIndex = Math.max(0, Math.min(destinationIndex, destinationSection.slides.length));
        destinationSection.slides.splice(boundedIndex, 0, movedSlide);

        return {
          ...currentBoard,
          sections,
        };
      });
    },
    [updateActiveBoard],
  );

  const handleDragEnd = useCallback(
    ({ active, over }) => {
      const activeData = active.data.current;
      setActiveDrag(null);

      if (!over || !activeData) {
        return;
      }

      if (activeData.type === 'section') {
        updateActiveBoard((currentBoard) => {
          const overTarget = resolveOverTarget(over, currentBoard);
          if (!overTarget) {
            return currentBoard;
          }

          const sourceIndex = currentBoard.sections.findIndex(
            (section) => section.id === activeData.sectionId,
          );

          let destinationSectionId = null;
          if (overTarget.type === 'section' || overTarget.type === 'lane' || overTarget.type === 'slide') {
            destinationSectionId = overTarget.sectionId;
          }

          if (!destinationSectionId || sourceIndex === -1) {
            return currentBoard;
          }

          const destinationIndex = currentBoard.sections.findIndex(
            (section) => section.id === destinationSectionId,
          );

          if (
            destinationIndex === -1 ||
            sourceIndex === destinationIndex
          ) {
            return currentBoard;
          }

          return {
            ...currentBoard,
            sections: arrayMove(currentBoard.sections, sourceIndex, destinationIndex),
          };
        });
        return;
      }

      if (activeData.type === 'slide') {
        updateActiveBoard((currentBoard) => {
          const overTarget = resolveOverTarget(over, currentBoard);
          if (!overTarget || (overTarget.type !== 'slide' && overTarget.type !== 'lane')) {
            return currentBoard;
          }

          const source = locateSlide(currentBoard, activeData.slideId);
          if (!source) {
            return currentBoard;
          }

          let destinationSectionId = source.sectionId;
          let destinationIndex = source.slideIndex;

          if (overTarget.type === 'slide') {
            const destination = locateSlide(currentBoard, overTarget.slideId);
            if (!destination) {
              return currentBoard;
            }
            destinationSectionId = destination.sectionId;
            destinationIndex = destination.slideIndex;
          } else if (overTarget.type === 'lane') {
            destinationSectionId = overTarget.sectionId;
            const destinationSection = currentBoard.sections.find(
              (section) => section.id === destinationSectionId,
            );
            destinationIndex = destinationSection ? destinationSection.slides.length - 1 : source.slideIndex;
          }

          const sections = currentBoard.sections.map((section) => ({
            ...section,
            slides: [...section.slides],
          }));

          const sourceSection = sections[source.sectionIndex];
          if (!sourceSection) {
            return currentBoard;
          }

          if (source.sectionId === destinationSectionId) {
            const finalIndex = Math.max(
              0,
              Math.min(destinationIndex, sourceSection.slides.length - 1),
            );

            if (source.slideIndex === finalIndex) {
              return currentBoard;
            }

            sourceSection.slides = arrayMove(
              sourceSection.slides,
              source.slideIndex,
              finalIndex,
            );

            return {
              ...currentBoard,
              sections,
            };
          }

          const destinationSection = sections.find(
            (section) => section.id === destinationSectionId,
          );
          if (!destinationSection) {
            return currentBoard;
          }

          const [movedSlide] = sourceSection.slides.splice(source.slideIndex, 1);
          if (!movedSlide) {
            return currentBoard;
          }

          const boundedIndex = Math.max(
            0,
            Math.min(destinationIndex, destinationSection.slides.length),
          );
          destinationSection.slides.splice(boundedIndex, 0, movedSlide);

          return {
            ...currentBoard,
            sections,
          };
        });
      }
    },
    [updateActiveBoard],
  );

  const exportToPptx = useCallback(() => {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';

    const slide = pptx.addSlide();

    const margin = 0.28;
    const sectionGap = 0.18;
    const sectionWidth = 2.15;
    const rightGap = 0.24;
    const rightWidth = EXPORT_WIDTH - margin * 2 - sectionWidth - rightGap;

    const lanePadX = 0.12;
    const lanePadY = 0.1;
    const cardGapX = 0.14;
    const cardGapY = 0.14;

    const maxSlides = Math.max(1, ...board.sections.map((section) => section.slides.length || 1));
    const columns = Math.min(4, Math.max(1, Math.ceil(Math.sqrt(maxSlides))));

    const rawHeights = board.sections.map((section) => {
      const slideCount = Math.max(1, section.slides.length);
      const rows = Math.ceil(slideCount / columns);
      const laneInnerHeight = rows * 0.82 + Math.max(0, rows - 1) * cardGapY;
      return Math.max(0.86, laneInnerHeight + lanePadY * 2);
    });

    const totalRawHeight =
      rawHeights.reduce((sum, height) => sum + height, 0) +
      sectionGap * Math.max(0, board.sections.length - 1);

    const availableHeight = EXPORT_HEIGHT - margin * 2;
    const computedScale = availableHeight / totalRawHeight;
    const scale = scaleToFit ? Math.min(1, computedScale) : 1;

    if (computedScale < 1) {
      console.warn(
        `Storyline content exceeds one slide at native size. Scale factor required: ${computedScale.toFixed(2)}.`,
      );
      if (!scaleToFit) {
        window.alert('Content exceeds one slide. Enable "Scale to Fit" for a guaranteed single-slide export.');
      }
    }

    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: EXPORT_WIDTH,
      h: EXPORT_HEIGHT,
      line: { color: 'FFFFFF', transparency: 100 },
      fill: { color: 'FFFFFF' },
    });

    slide.addText(board.title || 'Storyline Board', {
      x: margin,
      y: 0.08,
      w: EXPORT_WIDTH - margin * 2,
      h: 0.25,
      bold: true,
      fontFace: 'Calibri',
      fontSize: Math.max(11, 18 * scale),
      color: '0F172A',
      fit: 'resize',
    });

    const slideNumberById = buildSlideNumberMap(board.sections);

    let cursorY = margin + 0.22;

    board.sections.forEach((section, sectionIndex) => {
      const rawHeight = rawHeights[sectionIndex];
      const sectionHeight = rawHeight * scale;

      const laneX = margin + sectionWidth + rightGap;
      const laneY = cursorY;
      const laneW = rightWidth;
      const laneH = sectionHeight;

      slide.addShape(pptx.ShapeType.rect, {
        x: margin,
        y: cursorY,
        w: sectionWidth,
        h: sectionHeight,
        line: { color: SECTION_COLOR, pt: 1 },
        fill: { color: SECTION_COLOR },
      });

      slide.addText(`SECTION ${toSectionLabel(sectionIndex)}`, {
        x: margin + 0.1,
        y: cursorY + 0.06,
        w: sectionWidth - 0.2,
        h: 0.14,
        bold: true,
        fontFace: 'Calibri',
        fontSize: Math.max(6, 7.2 * scale),
        color: 'D7C8EE',
        fit: 'resize',
      });

      slide.addText(section.title || 'Section', {
        x: margin + 0.1,
        y: cursorY + 0.24,
        w: sectionWidth - 0.2,
        h: sectionHeight - 0.28,
        color: 'FFFFFF',
        bold: true,
        fontFace: 'Calibri',
        valign: 'top',
        fontSize: Math.max(8, 11 * scale),
        fit: 'resize',
      });

      const slideCount = Math.max(1, section.slides.length);
      const rows = Math.ceil(slideCount / columns);

      const cardWidth =
        (laneW - lanePadX * 2 - cardGapX * Math.max(0, columns - 1)) /
        Math.max(1, columns);
      const cardHeight =
        (laneH - lanePadY * 2 - cardGapY * Math.max(0, rows - 1)) /
        Math.max(1, rows);

      section.slides.forEach((deckSlide, slideIndex) => {
        const row = Math.floor(slideIndex / columns);
        const column = slideIndex % columns;

        const x = laneX + lanePadX + column * (cardWidth + cardGapX);
        const y = laneY + lanePadY + row * (cardHeight + cardGapY);

        slide.addShape(pptx.ShapeType.rect, {
          x,
          y,
          w: cardWidth,
          h: cardHeight,
          line: { color: 'C8D5E2', pt: 0.8 },
          fill: { color: CARD_BG },
        });

        const globalNumber = slideNumberById[deckSlide.id] ?? slideIndex + 1;

        slide.addText(String(globalNumber), {
          x: x + 0.05,
          y: y + 0.04,
          w: 0.18,
          h: 0.13,
          bold: true,
          fontFace: 'Calibri',
          color: '33475F',
          fontSize: Math.max(6, 7.2 * scale),
          fit: 'resize',
        });

        slide.addText((deckSlide.title || 'Slide').toUpperCase(), {
          x: x + 0.24,
          y: y + 0.04,
          w: cardWidth - 0.29,
          h: 0.18,
          bold: true,
          fontFace: 'Calibri',
          color: '1F3044',
          fontSize: Math.max(6.6, 8.2 * scale),
          fit: 'resize',
        });

        const bulletLines = deckSlide.bullets
          .map((bullet) => bullet.trim())
          .filter(Boolean)
          .map((bullet) => `• ${bullet}`)
          .join('\n');

        slide.addText(bulletLines || '•', {
          x: x + 0.06,
          y: y + 0.26,
          w: cardWidth - 0.12,
          h: Math.max(0.12, cardHeight - 0.3),
          fontFace: 'Calibri',
          color: '314A63',
          valign: 'top',
          fontSize: Math.max(6, 7.2 * scale),
          fit: 'shrink',
        });
      });

      cursorY += sectionHeight + sectionGap;
    });

    const fileName = safeFileName(board.title || 'storyline');
    pptx.writeFile({ fileName: `${fileName}.pptx` });
  }, [board, scaleToFit]);

  const sectionSortableIds = useMemo(
    () => board.sections.map((section) => `section:${section.id}`),
    [board.sections],
  );

  const activeSlide = useMemo(() => {
    if (activeDrag?.type !== 'slide') {
      return null;
    }

    return board.sections
      .flatMap((section) => section.slides)
      .find((slideCard) => slideCard.id === activeDrag.slideId) ?? null;
  }, [activeDrag, board.sections]);

  const activeSection = useMemo(() => {
    if (activeDrag?.type !== 'section') {
      return null;
    }

    return board.sections.find((section) => section.id === activeDrag.sectionId) ?? null;
  }, [activeDrag, board.sections]);

  return (
    <div className="ghost-app min-h-screen bg-[radial-gradient(circle_at_top_left,_#ffffff_0%,_#e8eef8_60%,_#dde7f6_100%)] px-6 py-8 text-slate-900">
      <div className="ghost-shell mx-auto max-w-[1400px]">
        <header className="ghost-header mb-6 rounded-2xl border border-slate-200/80 bg-white/90 p-6 shadow-sm backdrop-blur">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
            <div className="min-w-0 flex-1">
              <textarea
                ref={registerInputRef('story:title')}
                key={`story-title-${activeProject.id}`}
                rows={1}
                defaultValue={board.title}
                onInput={handleTextEditorInput}
                onChange={(event) => updateStoryTitle(event.target.value)}
                className="story-title text-editor w-full border-none bg-transparent text-4xl font-semibold tracking-tight text-slate-900 outline-none"
                placeholder="Storyline Title"
                aria-label="Storyline title"
              />
              <p className="story-hint mt-1 text-sm text-slate-500">
                Keyboard: Enter (title to bullet), Enter (new bullet), Backspace (delete empty bullet), Ctrl/Cmd + Enter (new slide)
              </p>
            </div>

            <div className="toolbar-actions flex flex-wrap items-center gap-2">
              <label className="toolbar-field flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-700">
                <FolderKanban size={16} />
                <span className="whitespace-nowrap">Board Switcher</span>
                <select
                  value={activeProject.id}
                  onChange={(event) => handleBoardSelection(event.target.value)}
                  className="toolbar-select max-w-[220px] bg-transparent text-sm outline-none"
                  aria-label="Board switcher"
                >
                  {projectStore.projects.map((project) => (
                    <option key={project.id} value={project.id}>
                      {project.name}
                    </option>
                  ))}
                  <option value="__new_board__">+ New Board</option>
                </select>
              </label>

              <label className="toolbar-field toolbar-toggle inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-700">
                <input
                  type="checkbox"
                  checked={scaleToFit}
                  onChange={(event) => setScaleToFit(event.target.checked)}
                  className="h-4 w-4"
                />
                Scale to Fit
              </label>

              <button
                type="button"
                onClick={exportBoardJson}
                className="toolbar-btn inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50"
              >
                <Download size={15} />
                Export JSON
              </button>

              <button
                type="button"
                onClick={openBoardImportDialog}
                className="toolbar-btn inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm font-semibold text-slate-700 hover:bg-slate-50"
              >
                <Upload size={15} />
                Import JSON
              </button>

              <input
                ref={boardImportInputRef}
                type="file"
                accept="application/json,.json"
                className="hidden"
                onChange={importBoardFromJson}
              />

              <button
                type="button"
                onClick={exportToPptx}
                className="toolbar-btn toolbar-btn-primary inline-flex items-center gap-2 rounded-lg bg-slate-900 px-4 py-2 text-sm font-semibold text-white hover:bg-slate-800"
              >
                <Download size={16} />
                Export PPTX
              </button>
            </div>
          </div>
        </header>

        <DndContext
          sensors={sensors}
          collisionDetection={closestCenter}
          onDragStart={handleDragStart}
          onDragOver={handleDragOver}
          onDragEnd={handleDragEnd}
          onDragCancel={handleDragCancel}
        >
          <SortableContext items={sectionSortableIds} strategy={verticalListSortingStrategy}>
            <div className="board-sections space-y-4">
              {board.sections.map((section, sectionIndex) => {
                const slideSortableIds = section.slides.map((slideCard) => `slide:${slideCard.id}`);

                return (
                  <Fragment key={section.id}>
                    <SortableSectionRow sectionId={section.id}>
                      {({ setNodeRef, style, dragHandleProps }) => (
                        <section
                          ref={setNodeRef}
                          style={style}
                          className="section-row group flex items-stretch gap-2"
                        >
                          <aside
                            className="section-anchor w-64 self-stretch rounded-xl p-3 text-white shadow-sm"
                            style={{ backgroundColor: SECTION_COLOR }}
                            {...dragHandleProps}
                          >
                            <div className="flex h-full flex-col">
                              <div className="mb-2 flex items-start justify-between gap-2">
                                <span className="section-kicker text-[11px] uppercase tracking-[0.16em] text-white/70">
                                  Section {toSectionLabel(sectionIndex)}
                                </span>
                                <div className="hover-actions section-actions flex gap-1 opacity-0 transition-opacity group-hover:opacity-100">
                                  <button
                                    type="button"
                                    onClick={() => duplicateSection(section.id)}
                                    className="rounded-md p-1 text-white/90 hover:bg-white/15"
                                    aria-label="Duplicate section"
                                  >
                                    <Copy size={14} />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => deleteSection(section.id)}
                                    className="rounded-md p-1 text-white/90 hover:bg-white/15"
                                    aria-label="Delete section"
                                  >
                                    <Trash2 size={14} />
                                  </button>
                                </div>
                              </div>

                              <textarea
                                ref={registerInputRef(getSectionTitleKey(section.id))}
                                key={`section-title-${section.id}`}
                                rows={1}
                                defaultValue={section.title}
                                onInput={handleTextEditorInput}
                                onChange={(event) => updateSectionTitle(section.id, event.target.value)}
                                className="section-title-input text-editor w-full border-none bg-transparent text-lg font-semibold leading-tight text-white outline-none placeholder:text-white/60"
                                placeholder="Section title"
                                aria-label="Section title"
                              />
                            </div>
                          </aside>

                          <SlideLaneDroppable sectionId={section.id}>
                            {({ setNodeRef: setLaneRef, isOver }) => (
                              <div
                                ref={setLaneRef}
                                className={`section-lane min-w-0 flex-1 rounded-xl border border-slate-200/80 bg-white/80 p-3 ${
                                  isOver ? 'is-over' : ''
                                }`}
                              >
                                <SortableContext
                                  items={slideSortableIds}
                                  strategy={rectSortingStrategy}
                                >
                                  <div className="slide-wrap flex flex-wrap items-start gap-3">
                                    {section.slides.map((slideCard, slideIndex) => (
                                      <SortableSlideSlot
                                        key={slideCard.id}
                                        slideId={slideCard.id}
                                        sectionId={section.id}
                                      >
                                        {({ setNodeRef: setSlideRef, style: slideStyle, dragHandleProps: slideDragProps, isDragging }) => (
                                          <div
                                            ref={setSlideRef}
                                            style={slideStyle}
                                            className={`slide-slot ${isDragging ? 'is-dragging' : ''}`}
                                          >
                                            <article
                                              {...slideDragProps}
                                              onBlurCapture={(event) => {
                                                const next = event.relatedTarget;
                                                if (!event.currentTarget.contains(next)) {
                                                  cleanupCardBullets(section.id, slideCard.id);
                                                }
                                              }}
                                              className="slide-card group/card border p-2"
                                              style={{ backgroundColor: CARD_BG, borderColor: CARD_BORDER }}
                                            >
                                              <div className="mb-1 flex items-start justify-between gap-1">
                                                <div className="slide-title-row flex min-w-0 flex-1 items-center gap-2">
                                                  <span className="slide-number">{slideNumberById[slideCard.id]}</span>
                                                  <textarea
                                                    ref={registerInputRef(getSlideTitleKey(slideCard.id))}
                                                    key={`slide-title-${slideCard.id}`}
                                                    rows={1}
                                                    defaultValue={slideCard.title}
                                                    onInput={handleTextEditorInput}
                                                    onChange={(event) =>
                                                      updateSlideTitle(section.id, slideCard.id, event.target.value)
                                                    }
                                                    onKeyDown={(event) =>
                                                      handleSlideTitleKeyDown(
                                                        event,
                                                        section.id,
                                                        slideCard.id,
                                                        slideCard.bullets.length,
                                                      )
                                                    }
                                                    className="slide-title-input text-editor w-full border-none bg-transparent font-semibold text-slate-900 outline-none placeholder:text-slate-400"
                                                    placeholder="Slide title"
                                                    aria-label="Slide title"
                                                  />
                                                </div>

                                                <div className="hover-actions slide-actions flex items-center gap-0.5 opacity-0 transition-opacity group-hover/card:opacity-100">
                                                  <button
                                                    type="button"
                                                    onClick={() => duplicateSlide(section.id, slideCard.id)}
                                                    className="icon-btn p-1 text-slate-500"
                                                    aria-label="Duplicate slide"
                                                  >
                                                    <Copy size={13} />
                                                  </button>
                                                  <button
                                                    type="button"
                                                    onClick={() => deleteSlide(section.id, slideCard.id)}
                                                    className="icon-btn p-1 text-slate-500"
                                                    aria-label="Delete slide"
                                                  >
                                                    <Trash2 size={13} />
                                                  </button>
                                                </div>
                                              </div>

                                              <div className="space-y-0.5">
                                                {slideCard.bullets.map((bullet, bulletIndex) => (
                                                  <div
                                                    key={`${slideCard.id}-bullet-${bulletIndex}`}
                                                    className="bullet-row flex items-start gap-1"
                                                  >
                                                    <span className="bullet-dot pt-1 text-slate-400">•</span>
                                                    <textarea
                                                      ref={registerInputRef(getBulletKey(slideCard.id, bulletIndex))}
                                                      key={`bullet-${slideCard.id}-${bulletIndex}`}
                                                      rows={1}
                                                      defaultValue={bullet}
                                                      onInput={handleTextEditorInput}
                                                      onChange={(event) =>
                                                        updateBulletText(
                                                          section.id,
                                                          slideCard.id,
                                                          bulletIndex,
                                                          event.target.value,
                                                        )
                                                      }
                                                      onKeyDown={(event) =>
                                                        handleBulletKeyDown(
                                                          event,
                                                          section.id,
                                                          slideCard.id,
                                                          bulletIndex,
                                                        )
                                                      }
                                                      className="bullet-input text-editor w-full border-none bg-transparent py-0.5 text-slate-700 outline-none placeholder:text-slate-400"
                                                      placeholder="Type bullet"
                                                      aria-label="Slide bullet"
                                                    />
                                                  </div>
                                                ))}
                                              </div>
                                            </article>

                                            <button
                                              type="button"
                                              onClick={() => insertSlideAt(section.id, slideIndex + 1)}
                                              className="slide-insert-btn"
                                              aria-label="Insert slide"
                                            >
                                              <Plus size={12} />
                                            </button>
                                          </div>
                                        )}
                                      </SortableSlideSlot>
                                    ))}
                                  </div>
                                </SortableContext>
                              </div>
                            )}
                          </SlideLaneDroppable>
                        </section>
                      )}
                    </SortableSectionRow>

                    <button
                      type="button"
                      onClick={() => insertSectionAt(sectionIndex + 1)}
                      className="section-insert-btn"
                      aria-label="Insert section"
                    >
                      <Plus size={12} />
                    </button>
                  </Fragment>
                );
              })}
            </div>
          </SortableContext>

          <DragOverlay>
            {activeSlide ? (
              <article
                className="slide-card drag-overlay-card border p-2"
                style={{ backgroundColor: CARD_BG, borderColor: CARD_BORDER, width: 'var(--card-w)' }}
              >
                <div className="slide-title-row flex min-w-0 items-center gap-2">
                  <span className="slide-number">{slideNumberById[activeSlide.id]}</span>
                  <div className="slide-title-input truncate">{activeSlide.title || 'Slide'}</div>
                </div>
              </article>
            ) : null}

            {activeSection ? (
              <aside
                className="section-anchor drag-overlay-section p-3 text-white"
                style={{ backgroundColor: SECTION_COLOR }}
              >
                <span className="section-kicker">Section</span>
                <div className="section-title-input mt-1">{activeSection.title || 'Section title'}</div>
              </aside>
            ) : null}
          </DragOverlay>
        </DndContext>
      </div>
    </div>
  );
}

export default App;














